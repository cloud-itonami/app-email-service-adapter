<script lang="ts">
  import { onMount } from "svelte";

  type ConnectionState = "not_connected" | "connected" | "syncing" | "error";
  type Connection = {
    connected?: boolean;
    display_name?: string;
    email?: string;
    expires_at?: string;
  };
  type SyncStatus = {
    last_synced_at?: string;
    emails_found?: number;
    emails_saved?: number;
    calendar_events_found?: number;
    calendar_events_saved?: number;
    error?: string;
  };
  type AuthSession = {
    accessJwt?: string;
    refreshJwt?: string;
    did?: string;
    handle?: string;
  };

  const SERVICE_BASE = "/xrpc/etzhayyim.outlook.v1.OutlookService";
  const REDIRECT_PATH = "/auth/callback";
  const AUTH_SESSION_KEY = "etzhayyim-auth-session";
  const AUTH_SIGNIN_URL = "https://authn.etzhayyim.com/sign-in";
  const OAUTH_STATE_KEY = "outlook_oauth_state_v1";
  const OAUTH_CODE_VERIFIER_KEY = "outlook_oauth_code_verifier_v1";

  let state: ConnectionState = "not_connected";
  let account = "not connected";
  let message = "Connect your Microsoft account to start mailbox sync.";
  let lastSync = "never";
  let busy = false;
  let errorText = "";
  let signedIn = false;
  let authActor = "";
  let connection: Connection | null = null;
  let syncStatus: SyncStatus | null = null;

  const timeline = [
    { title: "Auth", detail: "Microsoft OAuth 2.0 + PKCE" },
    { title: "Fetch", detail: "Mail + Calendar delta pull" },
    { title: "Index", detail: "etzhayyim knowledge graph update" },
  ];

  function callbackUrl(): string {
    return `${window.location.origin}${REDIRECT_PATH}`;
  }

  function loadAuthSession(): AuthSession | null {
    try {
      const raw = window.localStorage.getItem(AUTH_SESSION_KEY) ?? window.sessionStorage.getItem(AUTH_SESSION_KEY);
      if (!raw) return null;
      const parsed = JSON.parse(raw) as AuthSession;
      return parsed?.accessJwt ? parsed : null;
    } catch {
      return null;
    }
  }

  function saveAuthSession(session: AuthSession): void {
    if (!session?.accessJwt) return;
    const value = JSON.stringify(session);
    window.localStorage.setItem(AUTH_SESSION_KEY, value);
    window.sessionStorage.setItem(AUTH_SESSION_KEY, value);
  }

  function parseAuthCallbackSession(): AuthSession | null {
    const hash = window.location.hash.startsWith("#") ? window.location.hash.slice(1) : window.location.hash;
    const hashParams = new URLSearchParams(hash);
    const encodedAuth = hashParams.get("auth");
    if (encodedAuth) {
      try {
        const parsed = JSON.parse(decodeURIComponent(encodedAuth)) as AuthSession;
        if (parsed?.accessJwt) return parsed;
      } catch {
        // ignore malformed auth hash
      }
    }

    const search = new URLSearchParams(window.location.search);
    const accessJwt = search.get("accessJwt") || search.get("access_token");
    if (!accessJwt) return null;
    return {
      accessJwt,
      refreshJwt: search.get("refreshJwt") || search.get("refresh_token") || "",
      did: search.get("did") || search.get("sub") || "",
      handle: search.get("handle") || "",
    };
  }

  function effectiveToken(): string {
    return loadAuthSession()?.accessJwt ?? "";
  }

  function redirectToSignIn(): void {
    const redirect = encodeURIComponent(window.location.href);
    window.location.href = `${AUTH_SIGNIN_URL}?redirect_url=${redirect}`;
  }

  function loadOAuthProof(): { state: string; codeVerifier: string } {
    return {
      state: window.sessionStorage.getItem(OAUTH_STATE_KEY) ?? "",
      codeVerifier: window.sessionStorage.getItem(OAUTH_CODE_VERIFIER_KEY) ?? "",
    };
  }

  function saveOAuthProof(stateValue: string, codeVerifier: string): void {
    if (stateValue) {
      window.sessionStorage.setItem(OAUTH_STATE_KEY, stateValue);
    } else {
      window.sessionStorage.removeItem(OAUTH_STATE_KEY);
    }
    if (codeVerifier) {
      window.sessionStorage.setItem(OAUTH_CODE_VERIFIER_KEY, codeVerifier);
    } else {
      window.sessionStorage.removeItem(OAUTH_CODE_VERIFIER_KEY);
    }
  }

  function clearOAuthProof(): void {
    window.sessionStorage.removeItem(OAUTH_STATE_KEY);
    window.sessionStorage.removeItem(OAUTH_CODE_VERIFIER_KEY);
  }

  function parseOAuthError(search: URLSearchParams): string | null {
    const err = search.get("error");
    if (!err) return null;
    const desc = search.get("error_description");
    return desc ? `${err}: ${desc}` : err;
  }

  async function callApi<T = unknown>(method: string, body: Record<string, unknown> = {}): Promise<T> {
    const token = effectiveToken();

    const res = await fetch(`${SERVICE_BASE}.${method}`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        ...(token ? { Authorization: `Bearer ${token}` } : {}),
      },
      body: JSON.stringify(body),
    });

    const payload = await res.json().catch((error) => {
      console.warn('[silent-fail] outlook App.svelte: response json parse failed', error);
      return {};
    });
    if (!res.ok) {
      const msg = (payload as { message?: string; error?: string }).message
        ?? (payload as { message?: string; error?: string }).error
        ?? `Request failed: ${res.status}`;
      throw new Error(msg);
    }
    return payload as T;
  }

  async function refreshAuthStatus(): Promise<void> {
    try {
      const result = await callApi<{ signed_in?: boolean; actor_id?: string; user_id?: string }>("GetAuthStatus");
      signedIn = Boolean(result.signed_in);
      authActor = String(result.actor_id ?? result.user_id ?? "");
    } catch {
      signedIn = false;
      authActor = "";
    }
  }

  function applyConnection(conn: Connection | null): void {
    connection = conn;
    if (conn?.connected) {
      state = "connected";
      account = conn.email || conn.display_name || "Microsoft account";
      message = "Connected. You can start sync any time.";
      if (syncStatus?.last_synced_at) {
        lastSync = new Date(syncStatus.last_synced_at).toLocaleString();
      }
      return;
    }
    state = "not_connected";
    account = "not connected";
    message = "Connect your Microsoft account to start mailbox sync.";
    lastSync = "never";
  }

  async function refreshConnection(): Promise<void> {
    const result = await callApi<{ connection?: Connection }>("GetConnection");
    applyConnection(result.connection ?? null);
  }

  async function handleOAuthCallbackIfPresent(): Promise<boolean> {
    const search = new URLSearchParams(window.location.search);
    const oauthErr = parseOAuthError(search);
    if (oauthErr) {
      errorText = oauthErr;
      state = "error";
      clearOAuthProof();
      window.history.replaceState({}, "", "/");
      return true;
    }

    const code = search.get("code");
    if (!code) return false;

    busy = true;
    errorText = "";
    state = "syncing";
    message = "Completing OAuth callback...";

    try {
      const stateParam = search.get("state") ?? "";
      const proof = loadOAuthProof();
      if (proof.state && stateParam && proof.state !== stateParam) {
        throw new Error("OAuth state mismatch. Retry from Connect.");
      }

      const body: Record<string, unknown> = {
        code,
        redirect_uri: callbackUrl(),
      };
      if (proof.codeVerifier) body.code_verifier = proof.codeVerifier;
      if (stateParam) body.state = stateParam;

      const result = await callApi<{ connection?: Connection; sync?: SyncStatus }>("ExchangeCode", body);
      syncStatus = result.sync ?? null;
      if (syncStatus?.last_synced_at) {
        lastSync = new Date(syncStatus.last_synced_at).toLocaleString();
      }
      applyConnection(result.connection ?? null);
      message = "OAuth completed successfully.";
      clearOAuthProof();
      window.history.replaceState({}, "", "/");
      return true;
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      errorText = msg;
      state = "error";
      message = "OAuth callback failed.";
      return true;
    } finally {
      busy = false;
    }
  }

  async function startAuth(): Promise<void> {
    if (!signedIn) {
      redirectToSignIn();
      return;
    }
    errorText = "";
    busy = true;
    state = "syncing";
    message = "Opening Microsoft OAuth...";
    try {
      const result = await callApi<{ auth_url?: string; state?: string; code_verifier?: string }>("StartAuth", {
        redirect_uri: callbackUrl(),
      });
      saveOAuthProof(result.state ?? "", result.code_verifier ?? "");
      if (!result.auth_url) {
        throw new Error("StartAuth returned no auth_url.");
      }
      window.location.href = result.auth_url;
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      errorText = msg;
      state = "error";
      message = "Failed to start OAuth.";
      busy = false;
    }
  }

  async function syncNow(): Promise<void> {
    if (!connection?.connected || busy) return;
    errorText = "";
    busy = true;
    state = "syncing";
    message = "Sync in progress...";
    try {
      const result = await callApi<{ sync?: SyncStatus }>("SyncMailbox", { limit: 25 });
      syncStatus = result.sync ?? null;
      if (syncStatus?.last_synced_at) {
        lastSync = new Date(syncStatus.last_synced_at).toLocaleString();
      }
      state = "connected";
      message = "Sync complete. Mailbox and calendar are up to date.";
      if (syncStatus?.error) {
        errorText = `Sync error: ${syncStatus.error}`;
        state = "error";
      }
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      errorText = msg;
      state = "error";
      message = "Sync failed.";
    } finally {
      busy = false;
    }
  }

  async function disconnect(): Promise<void> {
    if (busy) return;
    errorText = "";
    busy = true;
    state = "syncing";
    message = "Disconnecting...";
    try {
      await callApi("Disconnect");
      syncStatus = null;
      applyConnection({ connected: false });
      message = "Disconnected. OAuth token revoked.";
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      errorText = msg;
      state = "error";
      message = "Disconnect failed.";
    } finally {
      busy = false;
    }
  }

  onMount(async () => {
    const callbackSession = parseAuthCallbackSession();
    if (callbackSession?.accessJwt) {
      saveAuthSession(callbackSession);
      const cleanUrl = `${window.location.origin}${window.location.pathname}`;
      window.history.replaceState({}, "", cleanUrl);
    }

    await refreshAuthStatus();
    if (!signedIn) {
      message = "Sign in with auth.etzhayyim.com to continue.";
      return;
    }
    if (authActor) {
      message = `Signed in as ${authActor}`;
    }

    const consumed = await handleOAuthCallbackIfPresent();
    if (!consumed) {
      try {
        busy = true;
        message = "Checking connection...";
        await refreshConnection();
        if (!connection?.connected) {
          message = "Authenticated. Outlook not connected yet.";
        }
      } catch (err) {
        const msg = err instanceof Error ? err.message : String(err);
        errorText = msg;
        state = "error";
      } finally {
        busy = false;
      }
    }
  });
</script>

<main class="shell">
  <section class="orb" aria-hidden="true"></section>
  <section class="panel">
    <header class="hero">
      <p class="eyebrow">etzhayyim Outlook Integration</p>
      <h1>Outlook OAuth Sync</h1>
      <p class="subtitle">Securely connect Microsoft account, then sync mail and calendar in one flow.</p>
    </header>

    <section class="tokenCard">
      <h2>Auth Session</h2>
      <p class="hint">Authentication is auto-detected from <code>auth.etzhayyim.com</code> session.</p>
      <div class="tokenRow">
        <div class="authSummary">
          <strong>{signedIn ? "Signed In" : "Signed Out"}</strong>
          <span>{authActor || "No active actor"}</span>
        </div>
        {#if signedIn}
          <button class="secondary" on:click={refreshAuthStatus} disabled={busy}>Refresh Auth</button>
        {:else}
          <button class="primary" on:click={redirectToSignIn} disabled={busy}>Sign In</button>
        {/if}
      </div>
    </section>

    <section class="statusCard">
      <div class="row">
        <span>Status</span>
        <strong class={`badge ${state}`}>
          {#if state === "not_connected"}Not Connected{/if}
          {#if state === "connected"}Connected{/if}
          {#if state === "syncing"}Syncing{/if}
          {#if state === "error"}Error{/if}
        </strong>
      </div>
      <div class="row">
        <span>Account</span>
        <strong>{account}</strong>
      </div>
      <div class="row">
        <span>Last Sync</span>
        <strong>{lastSync}</strong>
      </div>
      {#if connection?.expires_at}
        <div class="row">
          <span>Token Expires</span>
          <strong>{new Date(connection.expires_at).toLocaleString()}</strong>
        </div>
      {/if}

      <p class="message">{message}</p>

      <div class="actions">
        <button class="primary" on:click={startAuth} disabled={busy || !signedIn}>
          Connect Microsoft
        </button>
        <button class="secondary" on:click={syncNow} disabled={busy || !connection?.connected}>
          Sync Now
        </button>
        <button class="danger" on:click={disconnect} disabled={busy || !connection?.connected}>
          Disconnect
        </button>
      </div>

      {#if syncStatus}
        <div class="syncMeta">
          <p>Mail: found {syncStatus.emails_found ?? 0} / saved {syncStatus.emails_saved ?? 0}</p>
          <p>Calendar: found {syncStatus.calendar_events_found ?? 0} / saved {syncStatus.calendar_events_saved ?? 0}</p>
          {#if syncStatus.error}<p class="syncErr">Sync error: {syncStatus.error}</p>{/if}
        </div>
      {/if}

      {#if errorText}
        <p class="errorText">{errorText}</p>
      {/if}
    </section>

    <section class="flow">
      <h2>Sync Flow</h2>
      <ol>
        {#each timeline as step, idx}
          <li style={`--delay:${idx * 120}ms`}>
            <span class="dot"></span>
            <div>
              <strong>{step.title}</strong>
              <p>{step.detail}</p>
            </div>
          </li>
        {/each}
      </ol>
    </section>
  </section>
</main>

<style>
  :global(body) {
    --bg-a: #0c1b2e;
    --bg-b: #101a12;
    --ink: #e9f0f7;
    --muted: #a8b5c4;
    --card: #101820cc;
    --line: #263445;
    --accent: #3fb9ff;
    --ok: #32c98f;
    --warn: #f9b14a;
    --err: #f06a6a;
    margin: 0;
    min-height: 100vh;
    font-family: "Avenir Next", "Segoe UI", "Hiragino Kaku Gothic ProN", sans-serif;
    color: var(--ink);
    background:
      radial-gradient(1000px 440px at 90% -20%, #2f6fff44, transparent 60%),
      radial-gradient(900px 420px at 0% 100%, #00d08422, transparent 55%),
      linear-gradient(140deg, var(--bg-a), var(--bg-b));
  }

  .shell {
    min-height: 100vh;
    display: grid;
    place-items: center;
    padding: 2rem 1rem;
    position: relative;
    overflow: hidden;
  }

  .orb {
    position: absolute;
    width: 440px;
    aspect-ratio: 1;
    border-radius: 999px;
    background: radial-gradient(circle at 30% 30%, #8fe1ff66, #4f8dff11 68%);
    filter: blur(12px);
    right: -120px;
    top: -160px;
    animation: drift 9s ease-in-out infinite alternate;
  }

  .panel {
    width: min(880px, 100%);
    background: var(--card);
    border: 1px solid var(--line);
    border-radius: 24px;
    backdrop-filter: blur(8px);
    padding: 1.25rem;
    box-shadow: 0 18px 48px #0000004d;
    display: grid;
    gap: 1rem;
  }

  .hero h1 {
    margin: 0.15rem 0 0.4rem;
    font-size: clamp(1.6rem, 4vw, 2.3rem);
    letter-spacing: 0.02em;
  }

  .eyebrow {
    margin: 0;
    color: var(--accent);
    font-weight: 700;
    font-size: 0.82rem;
    text-transform: uppercase;
    letter-spacing: 0.12em;
  }

  .subtitle {
    margin: 0;
    color: var(--muted);
    line-height: 1.45;
  }

  .tokenCard,
  .statusCard,
  .flow {
    border: 1px solid var(--line);
    border-radius: 16px;
    padding: 1rem;
    background: #0c1218b3;
  }

  .tokenCard h2,
  .flow h2 {
    margin: 0 0 0.5rem;
    font-size: 1rem;
    color: #dce8f3;
  }

  .hint {
    margin: 0 0 0.6rem;
    color: var(--muted);
    font-size: 0.9rem;
  }

  .tokenRow {
    display: flex;
    gap: 0.6rem;
    flex-wrap: wrap;
  }

  .authSummary {
    flex: 1 1 320px;
    min-height: 44px;
    border: 1px solid #3a4c61;
    background: #101927;
    color: #e9f0f7;
    border-radius: 12px;
    padding: 0.55rem 0.75rem;
    font-size: 0.92rem;
    display: flex;
    flex-direction: column;
    justify-content: center;
    gap: 0.1rem;
  }

  .authSummary strong {
    color: #e9f0f7;
    font-size: 0.92rem;
  }

  .authSummary span {
    color: #9bb0c5;
    font-size: 0.82rem;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }

  .row {
    display: flex;
    justify-content: space-between;
    gap: 0.75rem;
    padding: 0.35rem 0;
    color: var(--muted);
    font-size: 0.94rem;
  }

  .row strong {
    color: var(--ink);
    font-weight: 600;
  }

  .badge {
    padding: 0.2rem 0.6rem;
    border-radius: 999px;
    border: 1px solid transparent;
    font-size: 0.82rem;
  }

  .badge.not_connected { color: var(--muted); border-color: #5f6a77; }
  .badge.connected { color: var(--ok); border-color: #32c98f77; }
  .badge.syncing { color: var(--warn); border-color: #f9b14a77; }
  .badge.error { color: var(--err); border-color: #f06a6a77; }

  .message {
    margin: 0.7rem 0 0;
    color: #d0dbea;
    font-size: 0.93rem;
  }

  .actions {
    margin-top: 0.95rem;
    display: flex;
    flex-wrap: wrap;
    gap: 0.6rem;
  }

  button {
    border: 0;
    border-radius: 12px;
    min-height: 44px;
    padding: 0.65rem 0.95rem;
    font-weight: 700;
    letter-spacing: 0.01em;
    cursor: pointer;
    transition: transform 120ms ease, opacity 120ms ease;
  }

  button:hover { transform: translateY(-1px); }
  button:disabled { opacity: 0.5; cursor: not-allowed; transform: none; }
  .primary { background: linear-gradient(120deg, #2d7dff, #35a8ff); color: white; }
  .secondary { background: #1f2a35; color: #d5e1ee; border: 1px solid #3a4c61; }
  .danger { background: #3a1f24; color: #ffb7b7; border: 1px solid #6d343f; }

  .syncMeta {
    margin-top: 0.8rem;
    font-size: 0.88rem;
    color: #bcd0e5;
  }

  .syncMeta p {
    margin: 0.2rem 0;
  }

  .syncErr,
  .errorText {
    color: #ff9b9b;
  }

  .errorText {
    margin-top: 0.7rem;
    font-size: 0.9rem;
  }

  .flow ol {
    list-style: none;
    margin: 0;
    padding: 0;
    display: grid;
    gap: 0.7rem;
  }

  .flow li {
    display: grid;
    grid-template-columns: 14px 1fr;
    gap: 0.6rem;
    opacity: 0;
    transform: translateY(6px);
    animation: reveal 420ms ease forwards;
    animation-delay: var(--delay);
  }

  .flow .dot {
    width: 10px;
    height: 10px;
    border-radius: 999px;
    margin-top: 0.28rem;
    background: linear-gradient(120deg, #6acbff, #7ce3b8);
    box-shadow: 0 0 0 4px #6acbff22;
  }

  .flow p {
    margin: 0.2rem 0 0;
    color: var(--muted);
    font-size: 0.9rem;
  }

  @keyframes reveal {
    to { opacity: 1; transform: translateY(0); }
  }

  @keyframes drift {
    from { transform: translate(0, 0) scale(1); }
    to { transform: translate(-20px, 18px) scale(1.05); }
  }
</style>
