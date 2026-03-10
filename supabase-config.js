(function () {
  "use strict";

  const SUPABASE_URL = "https://dlfekatfcdqgrroyrysn.supabase.co";
  const SUPABASE_ANON_KEY = "sb_publishable_xl85s-N_WV4ORMlmdiZPkA_4LPL3SdK";

  if (!window.supabase || !window.supabase.createClient) {
    console.error("Supabase JS não foi carregado.");
    return;
  }

  const supabase = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

  const TABLE_PROFILES = "profiles";
  const TABLE_PANEL_STATE = "panel_state";

  async function signIn(email, password) {
    const { data, error } = await supabase.auth.signInWithPassword({
      email,
      password
    });
    if (error) throw error;
    return data;
  }

  async function signOut() {
    const { error } = await supabase.auth.signOut();
    if (error) throw error;
  }

  async function getSession() {
    const { data, error } = await supabase.auth.getSession();
    if (error) throw error;
    return data.session;
  }

  async function getUser() {
    const { data, error } = await supabase.auth.getUser();
    if (error) throw error;
    return data.user;
  }

  async function getProfile() {
    const user = await getUser();
    if (!user) return null;

    const { data, error } = await supabase
      .from(TABLE_PROFILES)
      .select("id, email, role, full_name")
      .eq("id", user.id)
      .maybeSingle();

    if (error) throw error;
    return data;
  }

  async function requireAuth() {
    const session = await getSession();
    if (!session) {
      window.location.href = "login.html";
      return null;
    }
    return session;
  }

  async function requireAdmin() {
    const session = await requireAuth();
    if (!session) return null;

    const profile = await getProfile();
    if (!profile || profile.role !== "admin") {
      window.location.href = "index.html";
      return null;
    }

    return { session, profile };
  }

  async function loadPanelState() {
    const { data, error } = await supabase
      .from(TABLE_PANEL_STATE)
      .select("*")
      .eq("slug", "main")
      .maybeSingle();

    if (error) throw error;

    if (!data) {
      return {
        slug: "main",
        raw_rows: [],
        last_update: null,
        updated_by_email: null
      };
    }

    return data;
  }

  async function savePanelState(rawRows, updatedByEmail) {
    const payload = {
      slug: "main",
      raw_rows: rawRows,
      last_update: new Date().toISOString(),
      updated_by_email: updatedByEmail || null
    };

    const { data, error } = await supabase
      .from(TABLE_PANEL_STATE)
      .upsert(payload, { onConflict: "slug" })
      .select()
      .single();

    if (error) throw error;
    return data;
  }

  async function getCurrentUserWithProfile() {
    const session = await getSession();
    if (!session) return null;

    const user = session.user;
    const profile = await getProfile();

    return { user, profile };
  }

  function formatDateTimeBR(value) {
    if (!value) return "--";
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return "--";
    return date.toLocaleString("pt-BR");
  }

  window.CTSupabase = {
    client: supabase,
    signIn,
    signOut,
    getSession,
    getUser,
    getProfile,
    requireAuth,
    requireAdmin,
    loadPanelState,
    savePanelState,
    getCurrentUserWithProfile,
    formatDateTimeBR
  };
})();