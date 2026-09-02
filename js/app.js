(function () {
    "use strict";

    const CHANNEL_ID = "kontent-sp-picker";
    const GRAPH_RESOURCE = "https://graph.microsoft.com";
    const CLIENT_ID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    const TENANT_PATTERN = /^[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?$/i;

    const dom = {
        pickBtn: document.getElementById("btn-pick"),
        clearBtn: document.getElementById("btn-clear"),
        confirmBar: document.getElementById("confirm-clear"),
        confirmYes: document.getElementById("confirm-clear-yes"),
        confirmNo: document.getElementById("confirm-clear-no"),
        errorBanner: document.getElementById("error-banner"),
        errorText: document.getElementById("error-text"),
        errorDismiss: document.getElementById("error-dismiss"),
        fileDisplay: document.getElementById("file-display"),
        debugStatus: document.getElementById("debug-status"),
        container: document.getElementById("root-container"),
    };

    let config = null;
    let isDisabled = false;
    let msalInstance = null;
    let msalReady = false;
    let pickerWindow = null;
    let resizeObserver = null;

    // --- status / error UI -------------------------------------------------

    function setStatus(msg) {
        if (config && config.debug) {
            dom.debugStatus.textContent = msg;
            dom.debugStatus.classList.remove("hidden");
        }
        console.log("[SharePointPicker]", msg);
    }

    function showError(msg) {
        dom.errorText.textContent = msg;
        dom.errorBanner.classList.remove("hidden");
        console.error("[SharePointPicker]", msg);
    }

    function clearErrorBanner() {
        dom.errorBanner.classList.add("hidden");
        dom.errorText.textContent = "";
    }

    // --- config --------------------------------------------------------------

    function validateConfig(rawConfig) {
        const cfg = rawConfig || {};
        const errors = [];

        if (!cfg.clientId || !CLIENT_ID_PATTERN.test(cfg.clientId)) {
            errors.push("\"clientId\" is missing or is not a valid Azure AD application (client) ID.");
        }
        if (!cfg.sharePointTenant || !TENANT_PATTERN.test(cfg.sharePointTenant)) {
            errors.push("\"sharePointTenant\" is missing or is not a valid tenant name (e.g. \"contoso\" for contoso.sharepoint.com).");
        }
        if (cfg.selectionMode && cfg.selectionMode !== "single" && cfg.selectionMode !== "multiple") {
            errors.push("\"selectionMode\" must be either \"single\" or \"multiple\".");
        }

        if (errors.length) {
            return { valid: false, errors };
        }

        const teamSiteOrigin = `https://${cfg.sharePointTenant}.sharepoint.com`;
        const mySiteOrigin = `https://${cfg.sharePointTenant}-my.sharepoint.com`;

        return {
            valid: true,
            clientId: cfg.clientId,
            selectionMode: cfg.selectionMode === "single" ? "single" : "multiple",
            debug: !!cfg.debug,
            teamSiteOrigin,
            mySiteOrigin,
            pickerUrl: `${mySiteOrigin}/_layouts/15/FilePicker.aspx`,
            allowedOrigins: new Set([teamSiteOrigin, mySiteOrigin]),
        };
    }

    // --- MSAL ------------------------------------------------------------

    async function initMsal() {
        try {
            msalInstance = new msal.PublicClientApplication({
                auth: {
                    clientId: config.clientId,
                    authority: "https://login.microsoftonline.com/common",
                    redirectUri: window.location.origin + window.location.pathname,
                },
                cache: { cacheLocation: "sessionStorage" },
            });
            await msalInstance.initialize();
            msalReady = true;
            setStatus("MSAL ready | accounts: " + msalInstance.getAllAccounts().length);
        } catch (e) {
            showError("Microsoft authentication library failed to initialize: " + e.message);
        }
    }

    async function getTokenSilent(resource) {
        if (!msalReady) throw new Error("Authentication is not ready yet.");

        const resourceOrigin = new URL(resource).origin;
        const scopes = [`${resourceOrigin}/.default`];
        const accounts = msalInstance.getAllAccounts();
        if (!accounts.length) throw new Error("No signed-in Microsoft account.");

        try {
            const r = await msalInstance.acquireTokenSilent({ scopes, account: accounts[0] });
            return r.accessToken;
        } catch (err) {
            setStatus("Silent token acquisition failed for " + resourceOrigin + ", retrying with popup...");
            const r = await msalInstance.acquireTokenPopup({ scopes, account: accounts[0] });
            return r.accessToken;
        }
    }

    async function ensureAuthenticated() {
        const requiredResources = [config.mySiteOrigin, config.teamSiteOrigin, GRAPH_RESOURCE];
        let account = msalInstance.getAllAccounts()[0];

        if (!account) {
            const authRes = await msalInstance.loginPopup({ scopes: [`${config.mySiteOrigin}/.default`] });
            account = authRes.account;
        }

        for (const resource of requiredResources) {
            const scopes = [`${resource}/.default`];
            try {
                await msalInstance.acquireTokenSilent({ scopes, account });
            } catch (err) {
                await msalInstance.acquireTokenPopup({ scopes, account });
            }
        }
    }

    // --- rendering ---------------------------------------------------------

    function renderFiles(fileData) {
        const files = fileData ? (Array.isArray(fileData) ? fileData : [fileData]) : [];

        dom.fileDisplay.innerHTML = "";

        if (files.length === 0) {
            dom.fileDisplay.classList.add("hidden");
            dom.pickBtn.textContent = "Select SharePoint Files";
            dom.clearBtn.classList.add("hidden");
            hideConfirmClear();
            return;
        }

        files.forEach((file) => {
            const fileItem = document.createElement("div");
            fileItem.className = "file-item";

            const linkEl = document.createElement("a");
            linkEl.className = "file-link";
            linkEl.href = isSafeUrl(file.url) ? file.url : "#";
            linkEl.target = "_blank";
            linkEl.rel = "noopener noreferrer";
            linkEl.innerText = file.name || "View linked SharePoint file";

            const metaEl = document.createElement("div");
            metaEl.className = "file-meta";
            if (file.author === "Loading...") {
                metaEl.innerText = "Loading metadata...";
            } else {
                const dateStr = file.lastModified
                    ? new Date(file.lastModified).toLocaleDateString(undefined, {
                        year: "numeric",
                        month: "short",
                        day: "numeric",
                        hour: "2-digit",
                        minute: "2-digit",
                    })
                    : "Unknown date";
                metaEl.innerText = `Updated ${dateStr} by ${file.author || "Unknown author"}`;
            }

            fileItem.appendChild(linkEl);
            fileItem.appendChild(metaEl);
            dom.fileDisplay.appendChild(fileItem);
        });

        dom.fileDisplay.classList.remove("hidden");
        dom.pickBtn.textContent = "Change Files";
        if (!isDisabled) dom.clearBtn.classList.remove("hidden");
    }

    function isSafeUrl(url) {
        if (typeof url !== "string") return false;
        try {
            return new URL(url).protocol === "https:";
        } catch (e) {
            return false;
        }
    }

    function saveValue(fileData) {
        CustomElement.setValue(fileData ? JSON.stringify(fileData) : null);
        renderFiles(fileData);
    }

    function updateHeight() {
        if (dom.container) {
            CustomElement.setHeight(dom.container.getBoundingClientRect().height);
        }
    }

    // --- clear confirmation (no window.confirm - not available inside the
    // sandboxed custom element iframe) ---------------------------------------

    function showConfirmClear() {
        dom.confirmBar.classList.remove("hidden");
    }

    function hideConfirmClear() {
        dom.confirmBar.classList.add("hidden");
    }

    // --- picker launch -------------------------------------------------------

    function launchSharePointPicker() {
        const pickerOptions = {
            sdk: "8.0",
            action: "query",
            entry: { oneDrive: {} },
            authentication: { clientId: config.clientId },
            messaging: { origin: window.location.origin, channelId: CHANNEL_ID },
            selection: { mode: config.selectionMode },
        };

        const queryString = new URLSearchParams({ filePicker: JSON.stringify(pickerOptions) });
        const fullUrl = `${config.pickerUrl}?${queryString.toString()}`;
        pickerWindow = window.open(fullUrl, "SharePointPicker", "width=1000,height=700,menubar=no,toolbar=no");

        if (!pickerWindow) {
            showError("The file picker popup was blocked by your browser. Please allow popups for this site and try again.");
        }
    }

    // --- picker message/port handling ---------------------------------------

    function isAllowedOrigin(origin) {
        if (!config) return false;
        try {
            return config.allowedOrigins.has(new URL(origin).origin);
        } catch (e) {
            return false;
        }
    }

    async function handlePickerCommand(port, msg) {
        const cmdData = msg.data;
        setStatus("Picker command: " + cmdData.command + " (id:" + msg.id + ")");
        port.postMessage({ type: "acknowledge", id: msg.id });

        if (cmdData.command === "authenticate") {
            try {
                const token = await getTokenSilent(cmdData.resource);
                port.postMessage({ type: "result", id: msg.id, data: { result: "token", token } });
            } catch (e) {
                port.postMessage({ type: "error", id: msg.id, error: { code: "authFailed", message: e.message } });
            }
            return;
        }

        if (cmdData.command === "pick") {
            if (pickerWindow) pickerWindow.close();
            port.postMessage({ type: "result", id: msg.id, data: { result: "success" } });

            const selectedItems = cmdData.items || [];
            if (selectedItems.length === 0) return;

            const initialFiles = selectedItems.map((item) => ({
                name: item.name,
                url: item.webUrl || item["@content.downloadUrl"],
                id: item.id,
                author: "Loading...",
                lastModified: null,
            }));
            renderFiles(initialFiles);

            const finalFiles = await Promise.all(
                selectedItems.map(async (item) => {
                    let author = item.lastModifiedBy?.user?.displayName || item.createdBy?.user?.displayName;
                    let lastMod = item.lastModifiedDateTime;

                    if ((!author || !lastMod) && item.parentReference?.driveId) {
                        try {
                            const graphToken = await getTokenSilent(GRAPH_RESOURCE);
                            const url = `${GRAPH_RESOURCE}/v1.0/drives/${item.parentReference.driveId}/items/${item.id}?$select=lastModifiedBy,lastModifiedDateTime`;
                            const res = await fetch(url, { headers: { Authorization: `Bearer ${graphToken}` } });
                            if (res.ok) {
                                const data = await res.json();
                                author = author || data.lastModifiedBy?.user?.displayName;
                                lastMod = lastMod || data.lastModifiedDateTime;
                            }
                        } catch (e) {
                            console.error("[SharePointPicker] Metadata fetch failed", e);
                        }
                    }

                    return {
                        name: item.name,
                        url: item.webUrl || item["@content.downloadUrl"],
                        id: item.id,
                        author: author || "Unknown",
                        lastModified: lastMod || null,
                    };
                })
            );

            saveValue(finalFiles);
            return;
        }

        if (cmdData.command === "close") {
            if (pickerWindow) pickerWindow.close();
        }
    }

    window.addEventListener("message", (event) => {
        if (!isAllowedOrigin(event.origin)) return;

        const message = event.data;
        if (!message || message.type !== "initialize" || message.channelId !== CHANNEL_ID) return;

        const port = event.ports[0];
        setStatus("Picker port established.");

        port.addEventListener("message", (portEvent) => {
            const msg = portEvent.data;
            if (msg.type !== "command") return;
            handlePickerCommand(port, msg).catch((e) => showError("Unexpected picker error: " + e.message));
        });

        port.start();
        port.postMessage({ type: "activate" });
    });

    // --- event wiring ---------------------------------------------------------

    dom.errorDismiss.addEventListener("click", clearErrorBanner);

    dom.clearBtn.addEventListener("click", () => {
        clearErrorBanner();
        showConfirmClear();
    });

    dom.confirmNo.addEventListener("click", hideConfirmClear);

    dom.confirmYes.addEventListener("click", () => {
        hideConfirmClear();
        saveValue(null);
    });

    dom.pickBtn.addEventListener("click", async () => {
        clearErrorBanner();

        if (!msalReady) {
            showError("Authentication is still initializing. Please try again in a moment.");
            return;
        }

        dom.pickBtn.disabled = true;
        dom.pickBtn.textContent = "Authenticating...";

        try {
            await ensureAuthenticated();
            launchSharePointPicker();
        } catch (error) {
            showError("Microsoft authentication failed. Please ensure popups are allowed and try again.");
        } finally {
            dom.pickBtn.disabled = isDisabled;
            dom.pickBtn.textContent = dom.fileDisplay.classList.contains("hidden") ? "Select SharePoint Files" : "Change Files";
        }
    });

    // --- bootstrap -------------------------------------------------------------

    function applyDisabledState(disabled) {
        isDisabled = disabled;
        dom.pickBtn.disabled = disabled;
        if (disabled) {
            dom.clearBtn.classList.add("hidden");
            hideConfirmClear();
        } else if (!dom.fileDisplay.classList.contains("hidden")) {
            dom.clearBtn.classList.remove("hidden");
        }
    }

    CustomElement.init((element, context) => {
        const result = validateConfig(context.config);

        if (!result.valid) {
            showError(
                "This custom element is not configured correctly:\n" +
                result.errors.join("\n") +
                "\nSee the README for the expected configuration JSON."
            );
            dom.pickBtn.disabled = true;
            updateHeight();
            return;
        }

        config = result;
        if (config.debug) dom.debugStatus.classList.remove("hidden");
        setStatus("Origin: " + window.location.origin + " | MSAL: loading...");

        if (element.value) {
            try {
                renderFiles(JSON.parse(element.value));
            } catch (e) {
                console.error("[SharePointPicker] Failed to parse saved value", e);
            }
        }

        applyDisabledState(context.item.readOnly);
        initMsal();
    });

    if (typeof CustomElement.onDisabledChanged === "function") {
        CustomElement.onDisabledChanged(applyDisabledState);
    }

    resizeObserver = new ResizeObserver(() => updateHeight());
    resizeObserver.observe(dom.container);
})();
