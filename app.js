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
  lastTicketUrl: null,

  // v1.2.0 - Observadores
  observers: [],
  observerResults: [],
  observerActiveIndex: -1,
  observerTimer: null,
  observerAbort: null,

  // v1.2.0 - Modelos pré-prontos por categoria
  templateCache: new Map(),
  templateSeq: 0,
  appliedTemplate: null,
  pendingTemplate: null
};

const DESCRIPTION_MAX = 10000;

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
    setupObserverField();
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

  const description = byId("ticket-description").value.trim();
  const templateDescription = String(state.appliedTemplate?.description || "").trim();
  if (templateDescription && description === templateDescription) {
    setTicketAlert("Complete o modelo da descrição com as informações do seu chamado antes de enviar.", "error");
    byId("ticket-description").focus();
    return;
  }

  const formData = new FormData(form);
  for (const observer of state.observers) {
    formData.append("observerIds", String(observer.id));
  }
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
    const observerNames = (result.observers || []).map((item) => item.name).join(", ");
    byId("ticket-success-summary").textContent = observerNames
      ? `Chamado #${result.ticketId}: ${result.title} — Observadores: ${observerNames}`
      : `Chamado #${result.ticketId}: ${result.title}`;
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
  updateDescriptionCounter();
  setTicketAlert("", "info", true);
  filterAndPopulateCategories();

  state.templateSeq += 1;
  state.appliedTemplate = null;
  state.pendingTemplate = null;
  setTemplateHint("");

  state.observers = [];
  renderObserverChips();
  byId("ticket-observer-search").value = "";
  hideObserverResults();
}

/* =====================================================================
 * v1.2.0 - Modelos pré-prontos por categoria (Modelo de Chamado do GLPI)
 * ===================================================================== */

function updateDescriptionCounter() {
  const length = byId("ticket-description").value.length;
  byId("ticket-description-counter").textContent = `${length}/${DESCRIPTION_MAX}`;
}

function getSelectedCategory() {
  const id = Number(byId("ticket-category").value);
  return state.categoriesForType.find((category) => category.id === id) || null;
}

// Campo "intocado": vazio ou ainda igual ao texto do último modelo aplicado.
function isPristine(currentValue, appliedValue) {
  const current = String(currentValue || "").trim();
  return current === "" || (Boolean(appliedValue) && current === String(appliedValue).trim());
}

function setTemplateHint(message, showApply = false) {
  const hint = byId("ticket-template-hint");
  byId("ticket-template-hint-text").textContent = message || "";
  byId("ticket-template-apply").hidden = !showApply;
  hint.hidden = !message;
}

// Remove o texto do modelo anterior somente se o usuário não o editou.
function clearAppliedTemplateIfPristine() {
  const applied = state.appliedTemplate;
  if (!applied) return;

  const title = byId("ticket-title");
  const description = byId("ticket-description");

  if (applied.title && title.value.trim() === applied.title.trim()) title.value = "";
  if (applied.description && description.value.trim() === applied.description.trim()) description.value = "";

  state.appliedTemplate = null;
  updateDescriptionCounter();
}

function applyTemplate(template, force = false) {
  clearAppliedTemplateIfPristine();

  const title = byId("ticket-title");
  const description = byId("ticket-description");

  // Usuário já digitou em algum campo do modelo: não sobrescreve nada sem confirmação.
  const userTyped = (template.title && !isPristine(title.value)) ||
    (template.description && !isPristine(description.value));

  if (userTyped && !force) {
    state.pendingTemplate = template;
    setTemplateHint("Esta categoria possui um modelo pré-pronto. Deseja substituir o texto digitado?", true);
    return;
  }

  if (template.title) title.value = template.title;
  if (template.description) description.value = template.description.slice(0, DESCRIPTION_MAX);
  updateDescriptionCounter();

  state.appliedTemplate = template;
  state.pendingTemplate = null;
  setTemplateHint("Modelo da categoria aplicado. Complete as informações antes de enviar.");
}

async function onCategoryChange() {
  const seq = ++state.templateSeq;
  const type = byId("ticket-type").value;
  const category = getSelectedCategory();
  state.pendingTemplate = null;

  const hasTemplate = Boolean(category) &&
    (type === "incident" ? category.hasIncidentTemplate : category.hasRequestTemplate);

  if (!hasTemplate) {
    clearAppliedTemplateIfPristine();
    setTemplateHint("");
    return;
  }

  const cacheKey = `${type}:${category.id}`;
  setTemplateHint("Carregando modelo da categoria...");

  try {
    let template = state.templateCache.get(cacheKey);

    if (template === undefined) {
      const query = new URLSearchParams({ categoryId: String(category.id), type });
      const result = await apiRequest(`/templates?${query.toString()}`, { method: "GET" });
      template = result?.template || null;
      state.templateCache.set(cacheKey, template);
    }

    if (seq !== state.templateSeq) return; // categoria mudou durante a requisição

    if (!template) {
      clearAppliedTemplateIfPristine();
      setTemplateHint("");
      return;
    }

    applyTemplate(template);
  } catch (error) {
    if (seq !== state.templateSeq) return;
    console.warn("Falha ao carregar modelo da categoria", error);
    setTemplateHint("Não foi possível carregar o modelo desta categoria. Descreva o chamado normalmente.");
  }
}

function applyPendingTemplate() {
  if (!state.pendingTemplate) return;
  applyTemplate(state.pendingTemplate, true);
  byId("ticket-description").focus();
}

/* =====================================================================
 * v1.2.0 - Observadores (busca de usuários ativos no GLPI)
 * ===================================================================== */

function maxObservers() {
  const limit = Number(state.options?.limits?.maxObservers);
  return Number.isInteger(limit) && limit >= 0 ? limit : 5;
}

function setupObserverField() {
  const limit = maxObservers();
  byId("ticket-observers-group").hidden = limit === 0;
  byId("observer-limit-label").textContent = String(limit);
}

function hideObserverResults() {
  const list = byId("observer-results");
  list.hidden = true;
  list.replaceChildren();
  state.observerResults = [];
  state.observerActiveIndex = -1;
  const input = byId("ticket-observer-search");
  input.setAttribute("aria-expanded", "false");
  input.removeAttribute("aria-activedescendant");
}

function renderObserverMessage(message) {
  const list = byId("observer-results");
  const item = document.createElement("li");
  item.className = "observer-results-empty";
  item.textContent = message;
  list.replaceChildren(item);
  list.hidden = false;
  byId("ticket-observer-search").setAttribute("aria-expanded", "true");
}

function renderObserverResults() {
  const list = byId("observer-results");
  const input = byId("ticket-observer-search");

  if (state.observerResults.length === 0) {
    renderObserverMessage("Nenhum usuário ativo encontrado.");
    return;
  }

  const items = state.observerResults.map((user, index) => {
    const item = document.createElement("li");
    item.id = `observer-option-${index}`;
    item.className = "observer-option";
    item.setAttribute("role", "option");
    item.setAttribute("aria-selected", String(index === state.observerActiveIndex));

    const name = document.createElement("strong");
    name.textContent = user.name;
    const email = document.createElement("span");
    email.textContent = user.email || "";
    item.append(name, email);

    // mousedown evita que o blur do input feche a lista antes do clique.
    item.addEventListener("mousedown", (event) => event.preventDefault());
    item.addEventListener("click", () => addObserver(user));
    return item;
  });

  list.replaceChildren(...items);
  list.hidden = false;
  input.setAttribute("aria-expanded", "true");

  if (state.observerActiveIndex >= 0) {
    input.setAttribute("aria-activedescendant", `observer-option-${state.observerActiveIndex}`);
  } else {
    input.removeAttribute("aria-activedescendant");
  }
}

async function searchObservers(query) {
  state.observerAbort?.abort();
  const controller = new AbortController();
  state.observerAbort = controller;
  renderObserverMessage("Buscando usuários...");

  try {
    const result = await apiRequest(`/users/search?q=${encodeURIComponent(query)}`, {
      method: "GET",
      signal: controller.signal
    });

    if (controller.signal.aborted) return;

    const selected = new Set(state.observers.map((item) => item.id));
    state.observerResults = (result?.users || []).filter((user) => !selected.has(user.id));
    state.observerActiveIndex = state.observerResults.length > 0 ? 0 : -1;
    renderObserverResults();
  } catch (error) {
    if (error?.name === "AbortError" || controller.signal.aborted) return;
    console.warn("Falha na busca de observadores", error);
    renderObserverMessage(error.message || "Não foi possível buscar usuários.");
  }
}

function onObserverInput(event) {
  clearTimeout(state.observerTimer);
  const query = event.target.value.trim();

  if (query.length < 3) {
    state.observerAbort?.abort();
    hideObserverResults();
    return;
  }

  state.observerTimer = setTimeout(() => searchObservers(query), 300);
}

function onObserverKeydown(event) {
  const total = state.observerResults.length;

  if (event.key === "Enter") {
    // Nunca envia o formulário a partir do campo de busca.
    event.preventDefault();
    const user = state.observerResults[state.observerActiveIndex];
    if (user) addObserver(user);
    return;
  }

  if (event.key === "Escape") {
    event.stopPropagation();
    hideObserverResults();
    return;
  }

  if (total === 0 || (event.key !== "ArrowDown" && event.key !== "ArrowUp")) return;

  event.preventDefault();
  const step = event.key === "ArrowDown" ? 1 : -1;
  state.observerActiveIndex = (state.observerActiveIndex + step + total) % total;
  renderObserverResults();
  byId(`observer-option-${state.observerActiveIndex}`)?.scrollIntoView({ block: "nearest" });
}

function addObserver(user) {
  if (state.observers.some((item) => item.id === user.id)) return;

  if (state.observers.length >= maxObservers()) {
    renderObserverMessage(`Limite de ${maxObservers()} observador(es) atingido.`);
    return;
  }

  state.observers.push({ id: user.id, name: user.name, email: user.email || "" });
  renderObserverChips();

  const input = byId("ticket-observer-search");
  input.value = "";
  hideObserverResults();
  input.focus();
}

function removeObserver(userId) {
  state.observers = state.observers.filter((item) => item.id !== userId);
  renderObserverChips();
  byId("ticket-observer-search").focus();
}

function renderObserverChips() {
  const list = byId("observer-chips");
  const chips = state.observers.map((user) => {
    const chip = document.createElement("li");
    chip.className = "observer-chip";
    chip.title = user.email;

    const label = document.createElement("span");
    label.textContent = user.name;

    const remove = document.createElement("button");
    remove.type = "button";
    remove.textContent = "×";
    remove.setAttribute("aria-label", `Remover observador ${user.name}`);
    remove.addEventListener("click", () => removeObserver(user.id));

    chip.append(label, remove);
    return chip;
  });

  list.replaceChildren(...chips);
  list.hidden = chips.length === 0;
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
  byId("ticket-type").addEventListener("change", () => {
    filterAndPopulateCategories();
    onCategoryChange();
  });
  byId("ticket-category").addEventListener("change", onCategoryChange);
  byId("ticket-template-apply").addEventListener("click", applyPendingTemplate);
  byId("ticket-open-glpi").addEventListener("click", openLastTicket);
  byId("ticket-description").addEventListener("input", updateDescriptionCounter);

  const observerSearch = byId("ticket-observer-search");
  observerSearch.addEventListener("input", onObserverInput);
  observerSearch.addEventListener("keydown", onObserverKeydown);
  observerSearch.addEventListener("blur", () => setTimeout(hideObserverResults, 150));
  renderObserverChips();

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
