/**
 * Pré-validação de acesso ao Gerador de Relatório CCO.
 * Chama um endpoint autenticado do Hub; se a sessão for inválida, não libera o formulário.
 *
 * Personalização (atributos em document.body):
 *   data-auth-check-url — URL do check (padrão: "/api/relatorio")
 *   data-login-url      — Redirecionamento pós-401 (padrão: "/guest")
 *   data-home-url       — Link "Início" (padrão: "/home")
 */

const DEFAULT_CHECK_URL = "/api/relatorio";

function readBodyData(name, fallback) {
  const v = document.body?.getAttribute?.(name);
  return v && v.trim() ? v.trim() : fallback;
}

function setPreAuthVisible(visible) {
  const gate = document.getElementById("preAuthGate");
  if (!gate) return;
  gate.classList.toggle("hidden", !visible);
}

function setAppVisible(visible) {
  const app = document.getElementById("appContainer");
  if (!app) return;
  app.classList.toggle("hidden", !visible);
}

function updateGateUI({ message, showLogin, showRetry, showHome }) {
  const msg = document.getElementById("preAuthMessage");
  const login = document.getElementById("preAuthLoginLink");
  const retry = document.getElementById("preAuthRetry");
  const home = document.getElementById("preAuthHomeLink");
  if (msg) msg.textContent = message;
  if (login) login.classList.toggle("hidden", !showLogin);
  if (retry) retry.classList.toggle("hidden", !showRetry);
  if (home) home.classList.toggle("hidden", !showHome);
}

/**
 * Garante que o utilizador tem sessão válida antes de carregar a UI do relatório.
 * @returns {Promise<void>}
 */
export async function ensureCcoAccess() {
  setPreAuthVisible(true);
  setAppVisible(false);
  updateGateUI({
    message: "Verificando permissão para o relatório CCO…",
    showLogin: false,
    showRetry: false,
    showHome: false,
  });

  const checkUrl = readBodyData("data-auth-check-url", DEFAULT_CHECK_URL);
  const loginUrl = readBodyData("data-login-url", "/guest");
  const homeUrl = readBodyData("data-home-url", "/home");

  const loginLink = document.getElementById("preAuthLoginLink");
  const homeLink = document.getElementById("preAuthHomeLink");
  if (loginLink) loginLink.setAttribute("href", loginUrl);
  if (homeLink) homeLink.setAttribute("href", homeUrl);

  try {
    const res = await fetch(checkUrl, {
      method: "GET",
      credentials: "same-origin",
      headers: { Accept: "application/json" },
      cache: "no-store",
    });

    if (res.status === 401 || res.status === 403) {
      updateGateUI({
        message:
          "Sessão expirada ou sem permissão para o relatório CCO. Inicie sessão para continuar.",
        showLogin: true,
        showRetry: true,
        showHome: true,
      });
      throw new Error("unauthorized");
    }

    if (!res.ok) {
      updateGateUI({
        message: `Não foi possível validar o acesso (HTTP ${res.status}). Tente novamente.`,
        showLogin: false,
        showRetry: true,
        showHome: true,
      });
      throw { handled: true };
    }

    setPreAuthVisible(false);
    setAppVisible(true);
  } catch (e) {
    if (e && e.handled) throw e;
    if (e && e.message === "unauthorized") throw e;
    const isNetwork = e instanceof TypeError;
    updateGateUI({
      message: isNetwork
        ? "Falha de rede ao validar o acesso. Verifique a ligação e tente novamente."
        : "Não foi possível validar o acesso. Tente novamente.",
      showLogin: true,
      showRetry: true,
      showHome: true,
    });
    throw e;
  }
}

function wireRetry() {
  const retry = document.getElementById("preAuthRetry");
  if (!retry || retry.dataset.wired === "1") return;
  retry.dataset.wired = "1";
  retry.addEventListener("click", () => {
    ensureCcoAccess().catch(() => {});
  });
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", wireRetry);
} else {
  wireRetry();
}
