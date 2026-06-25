import { Route, Routes } from "react-router-dom";
import { Layout } from "./components/Layout";
import { HomePage } from "./pages/HomePage";
import { ScanDetailPage } from "./pages/ScanDetailPage";

export default function App() {
  return (
    <Layout>
      <Routes>
        <Route path="/" element={<HomePage />} />
        <Route path="/scans/:id" element={<ScanDetailPage />} />
      </Routes>
    </Layout>
  );
}
