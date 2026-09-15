"use strict";

const externalForms = Object.freeze({
  inovacao: "https://forms.office.com/Pages/ResponsePage.aspx?id=emONPyFz50qudewosdtpDCPJHkcW2CpMmyPiHvcb7dFUOUxBVkkzTExFVlZSWVNETjVSMkw5MDZJRi4u",
  "Antecipação de Pagamentos": "https://forms.office.com/Pages/ResponsePage.aspx?id=emONPyFz50qudewosdtpDDClf35qcbhFlPeWtNcJfuhUN0NOQlNZQ1lHMDI2WDBZNzhZSEtaVk1CQS4u",
  "Gestão de Conhecimento": "https://corpalloteamentos.sharepoint.com/:f:/r/sites/Corpal/Processos/GEST%C3%83O%20DE%20CONHECIMENTO?csf=1&web=1&e=nJYb7Y",
  "Cadastro de Fornecedor": "https://forms.office.com/r/z2xEgj6hC2?origin=lprLink",
  "Solicitação Controladoria": "https://forms.office.com/Pages/ResponsePage.aspx?id=emONPyFz50qudewosdtpDN6Q9G2Vru1GsBNOAniVhb9UQzBDWEZKSVFWQUVQNVpBVzY4RThYUU9IUy4u",
  "Manual da Marca": "https://corpalloteamentos.sharepoint.com/:f:/s/Corpal/IgAv5P8E0NBwQJMa2_nntyRtAT9e4t_8MOsCnDsx3sTQI8I?e=0wvo1P",
  viagem: "https://forms.cloud.microsoft/pages/responsepage.aspx?id=emONPyFz50qudewosdtpDN6Q9G2Vru1GsBNOAniVhb9URFUzTE1NRllWQlJON0Q1WkRGVkdNUFRERi4u",
  Manu_predial: "https://forms.cloud.microsoft/pages/responsepage.aspx?id=emONPyFz50qudewosdtpDN6Q9G2Vru1GsBNOAniVhb9UOEU4RDgwVEFMVTBVNzFFREtNWFBXMjE4SC4u&route=shorturl"
});

const state = {
  teamsInitialized: false,
  teamsContext: null,
  options: null,
  categoriesForType: [],
  lastTicketUrl: null
};

const config = window.CORPAL_APP_CONFIG;

function byId(id) {
  return document.getElementById(id);
}

function createRequestId() {
  if (globalThis.crypto?.randomUUID) {
    return globalThis.crypto.randomUUID();
  }

  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

async function initializeTeams() {
  if (state.teamsInitialized) {
    return;
  }

  if (!window.microsoftTeams?.app) {
    throw new Error("O Microsoft Teams SDK não foi carregado.");
  }

  await window.microsoftTeams.app.initialize();
  state.teamsContext = await window.microsoftTeams.app.getContext();
  state.teamsInitialized = true;
  window.microsoftTeams.app.notifyAppLoaded?.();
  renderTeamsUser();
}

function renderTeamsUser() {
  const box = byId("glpi-ticket-user");
  const user = state.teamsContext?.user;

  if (!box || !user) {
    return;
  }

  const displayName = user.displayName || "Usuário do Microsoft Teams";
  const loginHint = user.userPrincipalName || user.loginHint || "";
  box.textContent = loginHint ? `Solicitante: ${displayName} (${loginHint})` : `Solicitante: ${displayName}`;
  box.hidden = false;
}

async function getTeamsToken() {
  await initializeTeams();

  try {
    return await window.microsoftTeams.authentication.getAuthToken();
  } catch (error) {
    console.error("Falha ao obter token SSO do Teams", error);
    throw new Error("Não foi possível identificar seu usuário no Microsoft Teams. Feche e abra o aplicativo novamente.");
  }
}

async function apiRequest(path, options = {}) {
  const token = await getTeamsToken();
  const headers = new Headers(options.headers || {});
  headers.set("Authorization", `Bearer ${token}`);
  headers.set("X-Request-ID", createRequestId());

  const response = await fetch(`${config.apiBaseUrl}${path}`, {
    ...options,
    headers
  });

  const contentType = response.headers.get("content-type") || "";
  const payload = contentType.includes("application/json")
    ? await response.json()
    : { message: await response.text() };

  if (!response.ok) {
    const error = new Error(payload.message || payload.error || `Falha HTTP ${response.status}`);
    error.status = response.status;
    error.details = payload.details;
    throw error;
  }

  return payload;
}

function showOnly(sectionId) {
  const sections = ["icons", "form-container", "glpi-ticket-container"];

  for (const id of sections) {
    const element = byId(id);
    if (!element) continue;

    const visible = id === sectionId;
    element.hidden = !visible;
    if (id === "icons") {
      element.style.display = visible ? "flex" : "none";
    }
  }
}

function openExternal(url) {
  const opened = window.open(url, "_blank", "noopener,noreferrer");
  if (opened) {
    opened.opener = null;
  }
}

async function loadForm(type) {
  if (type === "chamados TI") {
    await glpiTicketOpen();
    return;
  }

  const url = externalForms[type];
  if (url) {
    openExternal(url);
  }
}

async function glpiTicketOpen() {
  showOnly("glpi-ticket-container");
  setTicketAlert("", "info", true);
  byId("glpi-ticket-success").hidden = true;
  byId("glpi-ticket-form").hidden = false;

  try {
    await initializeTeams();
    await loadTicketOptions();
  } catch (error) {
    setTicketAlert(error.message, "error");
  }
}

async function loadTicketOptions() {
  setOptionsLoading(true);

  try {
    state.options = await apiRequest("/options", { method: "GET" });

    const serverLimit = Number(state.options?.limits?.maxAttachmentMb);
    if (Number.isFinite(serverLimit) && serverLimit > 0) {
      byId("attachment-limit-label").textContent = `${serverLimit} MB`;
    }

    populateLocations(state.options.locations || []);
    filterAndPopulateCategories();
  } finally {
    setOptionsLoading(false);
  }
}

function setOptionsLoading(isLoading) {
  const category = byId("ticket-category");
  const location = byId("ticket-location");

  category.disabled = isLoading;
  location.disabled = isLoading;
}

function filterAndPopulateCategories() {
  const type = byId("ticket-type").value;
  const categories = (state.options?.categories || []).filter((category) => {
    return type === "incident" ? category.isIncident : category.isRequest;
  });

  state.categoriesForType = categories;
  const select = byId("ticket-category");
  select.replaceChildren(new Option("Selecione uma categoria", ""));

  for (const category of categories) {
    select.add(new Option(category.name, String(category.id)));
  }

  select.disabled = false;
}

function populateLocations(locations) {
  const select = byId("ticket-location");
  select.replaceChildren(new Option("Selecione uma localização", ""));

  for (const location of locations) {
    select.add(new Option(location.name, String(location.id)));
  }

  select.disabled = false;
}

function setTicketAlert(message, type = "error", hide = false) {
  const alert = byId("glpi-ticket-alert");
  alert.hidden = hide || !message;
  alert.textContent = message || "";
  alert.className = `ticket-alert ticket-alert-${type}`;
}

function setSubmitting(isSubmitting) {
  byId("ticket-submit").disabled = isSubmitting;
  byId("ticket-submit-label").textContent = isSubmitting ? "Enviando..." : "Abrir chamado";
  byId("ticket-submit-spinner").hidden = !isSubmitting;
}

async function submitTicket(event) {
  event.preventDefault();
  setTicketAlert("", "info", true);

  const form = event.currentTarget;

  if (!form.reportValidity()) {
    return;
  }

  const file = byId("ticket-attachment").files[0];
  const maxMb = Number(state.options?.limits?.maxAttachmentMb || config.maxAttachmentMb || 10);

  if (file && file.size > maxMb * 1024 * 1024) {
    setTicketAlert(`O anexo excede o limite de ${maxMb} MB.`, "error");
    return;
  }

  const formData = new FormData(form);
  const idempotencyKey = createRequestId();

  setSubmitting(true);

  try {
    const result = await apiRequest("/tickets", {
      method: "POST",
      body: formData,
      headers: {
        "Idempotency-Key": idempotencyKey
      }
    });

    state.lastTicketUrl = result.ticketUrl;
    byId("ticket-success-summary").textContent = `Chamado #${result.ticketId}: ${result.title}`;
    byId("glpi-ticket-form").hidden = true;
    byId("glpi-ticket-success").hidden = false;

    if (Array.isArray(result.warnings) && result.warnings.length > 0) {
      setTicketAlert(result.warnings.join(" "), "warning");
    }
  } catch (error) {
    console.error("Falha ao criar chamado", error);
    setTicketAlert(error.message || "Não foi possível criar o chamado.", "error");
  } finally {
    setSubmitting(false);
  }
}

function glpiTicketReset() {
  const form = byId("glpi-ticket-form");
  form.reset();
  byId("ticket-type").value = "incident";
  byId("ticket-description-counter").textContent = "0/10000";
  setTicketAlert("", "info", true);
  filterAndPopulateCategories();
}

function glpiTicketNew() {
  glpiTicketReset();
  byId("glpi-ticket-success").hidden = true;
  byId("glpi-ticket-form").hidden = false;
  byId("ticket-title").focus();
}

function openLastTicket() {
  if (!state.lastTicketUrl) return;

  if (window.microsoftTeams?.app?.openLink) {
    window.microsoftTeams.app.openLink(state.lastTicketUrl).catch(() => openExternal(state.lastTicketUrl));
  } else {
    openExternal(state.lastTicketUrl);
  }
}

function voltarInicio() {
  const frame = byId("form-frame");
  frame.src = "about:blank";
  showOnly("icons");
  setTicketAlert("", "info", true);
}

function openRHModal() {
  byId("rh-modal").style.display = "flex";
}

function closeRHModal(event) {
  const modal = byId("rh-modal");
  if (!event || event.target === modal) {
    modal.style.display = "none";
  }
}

function abrirRH(opcao) {
  closeRHModal();

  if (opcao === "play") {
    openExternal("https://play.google.com/store/apps/details?id=com.wiipo&pcampaignid=web_share");
  } else if (opcao === "apple") {
    openExternal("https://apps.apple.com/br/app/wiipo/id1522728600");
  }
}

document.addEventListener("DOMContentLoaded", () => {
  byId("glpi-ticket-form").addEventListener("submit", submitTicket);
  byId("ticket-type").addEventListener("change", filterAndPopulateCategories);
  byId("ticket-open-glpi").addEventListener("click", openLastTicket);
  byId("ticket-description").addEventListener("input", (event) => {
    byId("ticket-description-counter").textContent = `${event.target.value.length}/10000`;
  });

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      closeRHModal();
    }
  });

  initializeTeams().catch(() => {
    // Os formulários externos continuam funcionando fora do Teams.
  });
});

Object.assign(window, {
  loadForm,
  voltarInicio,
  openRHModal,
  closeRHModal,
  abrirRH,
  glpiTicketOpen,
  glpiTicketReset,
  glpiTicketNew
});
