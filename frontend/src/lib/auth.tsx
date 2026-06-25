import {
  createContext,
  useCallback,
  useContext,
  useEffect,
  useRef,
  useState,
  type ReactNode,
} from "react";
import { api, setAuthToken } from "./api";
import type { AppConfig, AppUser } from "./types";

interface GoogleCredentialResponse {
  credential: string;
}
declare global {
  interface Window {
    google?: {
      accounts: {
        id: {
          initialize: (cfg: {
            client_id: string;
            callback: (r: GoogleCredentialResponse) => void;
            auto_select?: boolean;
          }) => void;
          renderButton: (el: HTMLElement, opts: Record<string, unknown>) => void;
          disableAutoSelect: () => void;
        };
      };
    };
  }
}

const TOKEN_KEY = "decoder_id_token";

interface AuthState {
  ready: boolean;
  config: AppConfig | null;
  user: AppUser | null;
  authError: string | null;
  signOut: () => void;
  renderButton: (el: HTMLElement) => void;
}

const AuthContext = createContext<AuthState | null>(null);

function waitForGoogle(timeoutMs = 5000): Promise<Window["google"] | null> {
  return new Promise((resolve) => {
    const start = Date.now();
    const tick = () => {
      if (window.google) return resolve(window.google);
      if (Date.now() - start > timeoutMs) return resolve(null);
      setTimeout(tick, 100);
    };
    tick();
  });
}

export function AuthProvider({ children }: { children: ReactNode }) {
  const [ready, setReady] = useState(false);
  const [config, setConfig] = useState<AppConfig | null>(null);
  const [user, setUser] = useState<AppUser | null>(null);
  const [authError, setAuthError] = useState<string | null>(null);
  const initialized = useRef(false);

  const handleCredential = useCallback(async (resp: GoogleCredentialResponse) => {
    setAuthToken(resp.credential);
    localStorage.setItem(TOKEN_KEY, resp.credential);
    try {
      const me = await api.getMe();
      setAuthError(null);
      setUser(me);
    } catch (e) {
      // The token was valid but the backend rejected this account (e.g. wrong
      // domain). Surface the reason instead of silently bouncing to sign-in.
      setAuthToken(null);
      localStorage.removeItem(TOKEN_KEY);
      setUser(null);
      setAuthError(e instanceof Error ? e.message : "Sign-in failed.");
      window.google?.accounts.id.disableAutoSelect();
    }
  }, []);

  useEffect(() => {
    (async () => {
      const cfg = await api.getConfig();
      setConfig(cfg);

      if (!cfg.authEnabled) {
        try {
          setUser(await api.getMe());
        } catch {
          /* local mode still returns a synthetic user */
        }
        setReady(true);
        return;
      }

      // Restore a stored token if it's still valid.
      const stored = localStorage.getItem(TOKEN_KEY);
      if (stored) {
        setAuthToken(stored);
        try {
          setUser(await api.getMe());
        } catch {
          setAuthToken(null);
          localStorage.removeItem(TOKEN_KEY);
        }
      }

      const google = await waitForGoogle();
      if (google && cfg.oauthClientId && !initialized.current) {
        google.accounts.id.initialize({
          client_id: cfg.oauthClientId,
          callback: handleCredential,
          auto_select: true,
        });
        initialized.current = true;
      }
      setReady(true);
    })();
  }, [handleCredential]);

  const signOut = useCallback(() => {
    setUser(null);
    setAuthError(null);
    setAuthToken(null);
    localStorage.removeItem(TOKEN_KEY);
    window.google?.accounts.id.disableAutoSelect();
  }, []);

  const renderButton = useCallback((el: HTMLElement) => {
    if (window.google && config?.oauthClientId) {
      window.google.accounts.id.renderButton(el, {
        theme: "outline",
        size: "large",
        text: "signin_with",
        shape: "pill",
      });
    }
  }, [config]);

  return (
    <AuthContext.Provider value={{ ready, config, user, authError, signOut, renderButton }}>
      {children}
    </AuthContext.Provider>
  );
}

export function useAuth(): AuthState {
  const ctx = useContext(AuthContext);
  if (!ctx) throw new Error("useAuth must be used within AuthProvider");
  return ctx;
}
