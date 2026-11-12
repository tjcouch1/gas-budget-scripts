import { StrictMode, useEffect, useState } from "react";
import { createRoot } from "react-dom/client";
import { registerSW } from "virtual:pwa-register";
import { App } from "./App.tsx";
import { createServices } from "./services.ts";
import "./index.css";

const services = createServices();

function Root() {
  const [updateAvailable, setUpdateAvailable] = useState(false);
  const [updateSW, setUpdateSW] = useState<
    ((reload?: boolean) => Promise<void>) | null
  >(null);

  useEffect(() => {
    const update = registerSW({
      onNeedRefresh: () => setUpdateAvailable(true),
    });
    setUpdateSW(() => update);
  }, []);

  return (
    <App
      services={services}
      updateAvailable={updateAvailable}
      onUpdate={() => void updateSW?.(true)}
    />
  );
}

createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <Root />
  </StrictMode>,
);
