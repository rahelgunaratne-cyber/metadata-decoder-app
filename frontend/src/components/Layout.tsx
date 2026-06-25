import { useEffect, useRef, type ReactNode } from "react";
import { Link } from "react-router-dom";
import { useAuth } from "../lib/auth";
import { Button, Card, Spinner } from "./ui";

function Brand() {
  return (
    <Link to="/" className="flex items-center gap-2.5">
      <span className="grid h-8 w-8 place-items-center rounded-lg bg-navy text-base font-bold text-white">
        M
      </span>
      <div className="leading-tight">
        <div className="text-[15px] font-bold text-navy">Metadata Decoder</div>
        <div className="text-[11px] text-muted">Catch mistakes before ingestion</div>
      </div>
    </Link>
  );
}

function SignInScreen() {
  const { renderButton, config, authError } = useAuth();
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (ref.current) renderButton(ref.current);
  }, [renderButton]);

  return (
    <div className="grid min-h-screen place-items-center px-4">
      <Card className="w-full max-w-md p-8 text-center">
        <div className="mx-auto mb-4 grid h-12 w-12 place-items-center rounded-xl bg-navy text-xl font-bold text-white">
          M
        </div>
        <h1 className="text-xl font-bold text-navy">Metadata Decoder</h1>
        <p className="mt-1 text-sm text-muted">
          Sign in with your{" "}
          <span className="font-medium">@{config?.allowedDomain}</span> Google account to
          continue.
        </p>
        {authError && (
          <div className="mt-4 rounded-lg border border-artist/60 bg-artist/30 px-3 py-2 text-sm text-artist-ink">
            {authError}
          </div>
        )}
        <div className="mt-6 flex justify-center" ref={ref} />
      </Card>
    </div>
  );
}

export function Layout({ children }: { children: ReactNode }) {
  const { ready, config, user, signOut } = useAuth();

  if (!ready) {
    return (
      <div className="grid min-h-screen place-items-center">
        <Spinner className="h-6 w-6" />
      </div>
    );
  }

  if (config?.authEnabled && !user) {
    return <SignInScreen />;
  }

  return (
    <div className="min-h-screen">
      <header className="border-b border-slate-200 bg-white">
        <div className="mx-auto flex max-w-6xl items-center justify-between px-6 py-3">
          <Brand />
          {user && (
            <div className="flex items-center gap-3">
              {user.picture ? (
                <img src={user.picture} alt="" className="h-8 w-8 rounded-full" />
              ) : null}
              <div className="hidden text-right sm:block">
                <div className="text-sm font-medium text-ink">{user.name}</div>
                <div className="text-xs text-muted">{user.email}</div>
              </div>
              {config?.authEnabled && (
                <Button variant="ghost" onClick={signOut}>
                  Sign out
                </Button>
              )}
            </div>
          )}
        </div>
      </header>
      <main className="mx-auto max-w-6xl px-6 py-8">{children}</main>
    </div>
  );
}
