// Document Study App
// ============================================================

(async () => {

    const msalConfig = {
        auth: {
            clientId: "895b11a2-7e6e-4068-80f6-ce6bcf8c124e",
            authority: "https://login.microsoftonline.com/f2338aa2-defa-4cb8-815e-7b85c7271fec",
            redirectUri: window.location.origin
        }
    };

    const msalInstance = new msal.PublicClientApplication(msalConfig);

    // ✅ IMPORTANT: process redirect login (required for My Apps)
    await msalInstance.handleRedirectPromise();

    // ── AUTH SETUP ─────────────────────────────

    function getAccount() {
        const accounts = msalInstance.getAllAccounts();
        return accounts.length > 0 ? accounts[0] : null;
    }

    async function initAuth() {
        let account = getAccount();
        if (account) return account;

        try {
            // ✅ Force silent SSO via redirect
            await msalInstance.loginRedirect({
                scopes: ["User.Read"],
                prompt: "none"
            });

            return null; // browser will redirect
        } catch (err) {
            console.log("Silent redirect failed:", err);

            // fallback → show login screen
            showLoginScreen();
            return null;
        }
    }

    function showLoginScreen() {
        document.body.innerHTML = `
        <div style="text-align:center;margin-top:100px;">
            <h2>Sign in required</h2>
            <button id="loginBtn">Login with Microsoft</button>
        </div>
        `;

        document.getElementById("loginBtn").onclick = () => {
            msalInstance.loginRedirect({ scopes: ["User.Read"] });
        };

        // ✅ show UI AFTER replacing content
        document.body.style.display = "block";
    }

    function addAuthUI(account) {
        const header = document.querySelector("header");

        const userDiv = document.createElement("div");
        userDiv.style.marginTop = "10px";
        userDiv.innerHTML = `
            <div style="font-size: 0.9rem;">
                ${account.username}
            </div>
        `;

        header.appendChild(userDiv);
    }

    // ✅ AUTH FLOW
    const user = await initAuth();
    if (!user) return;

    addAuthUI(user);

    // ✅ remove loading + show app
    document.getElementById("loading")?.remove();
    document.body.style.display = "block";

    // ── VIEW FUNCTIONS ─────────────────────────

    function showListView() {
        document.getElementById("resultsView").style.display = "block";
        document.getElementById("readerView").style.display = "none";
        document.getElementById("controls").style.display = "block";
    }

    function showReaderView() {
        document.getElementById("resultsView").style.display = "none";
        document.getElementById("readerView").style.display = "block";
        document.getElementById("controls").style.display = "none";
    }

    function parseFilename(raw) {
        const cleaned = raw.replace(/\.[a-zA-Z0-9]+$/, "").trim();
        const parts = cleaned.split(" - ").map(p => p.trim());

        return {
            title: parts[0] || "Untitled",
            author: parts[1] || "Unknown",
            century: parts[2] || "?"
        };
    }

    // ── Fetch documents ──
    const FILENAMES = await fetch('/docs/files.json')
        .then(res => res.json());

    const documents = FILENAMES.map(filename => {
        const meta = parseFilename(filename);

        return {
            filename,
            title: meta.title,
            author: meta.author,
            century: meta.century
        };
    });

    // ── Populate filter ──
    function populateCenturyFilter() {
        const centuries = [...new Set(
            documents.map(d => d.century?.trim()).filter(Boolean)
        )].sort();

        filterCentury.innerHTML = `<option value="">All centuries</option>`;

        centuries.forEach(c => {
            const option = document.createElement("option");
            option.value = c;
            option.textContent = c;
            filterCentury.appendChild(option);
        });
    }

    // ── DOM ──
    const searchInput = document.getElementById("searchInput");
    const sortSelect = document.getElementById("sortSelect");
    const filterCentury = document.getElementById("filterCentury");
    const clearBtn = document.getElementById("clearBtn");

    const results = document.getElementById("results");
    const contentContainer = document.getElementById("contentContainer");

    // ── Render list ──
    function render(docs) {
        if (!docs.length) {
            results.innerHTML = `<div class="no-results">No documents found.</div>`;
            return;
        }

        results.innerHTML = docs.map(doc => `
            <div class="doc-card" data-filename="${encodeURIComponent(doc.filename)}">
                <div class="title">${escapeHtml(doc.title)}</div>
                <div class="meta">
                    <span>✍️ ${escapeHtml(doc.author)}</span>
                    <span>📅 ${escapeHtml(doc.century)}</span>
                </div>
            </div>
        `).join("");

        document.querySelectorAll(".doc-card").forEach(card => {
            card.addEventListener("click", async () => {
                const filename = decodeURIComponent(card.dataset.filename);
                showReaderView();
                await loadDocumentContent(filename);
            });
        });
    }

    // ── Load document ──
    async function loadDocumentContent(filename) {
        const response = await fetch(`/docs/${encodeURIComponent(filename)}`);
        const content = await response.text();

        const doc = documents.find(d => d.filename === filename);

        contentContainer.innerHTML = `
            <button id="backBtn">Back</button>
            <h2>${escapeHtml(doc?.title || filename)}</h2>
            <p>${escapeHtml(doc?.author || "")} ${escapeHtml(doc?.century || "")}</p>
            <pre>${escapeHtml(content)}</pre>
        `;

        document.getElementById("backBtn").onclick = () => {
            contentContainer.innerHTML = "";
            showListView();
        };
    }

    // ── Escape HTML (fixed) ──
    function escapeHtml(str) {
        return str
            .replaceAll("&", "&amp;")
            .replaceAll("<", "&lt;")
            .replaceAll(">", "&gt;")
            .replaceAll('"', "&quot;")
            .replaceAll("'", "&#039;");
    }

    // ── Filtering + sorting ──
    function updateResults() {
        const q = searchInput.value.toLowerCase().trim();
        const centuryFilterVal = filterCentury.value;
        const sortKey = sortSelect.value;

        let filtered = documents.filter(d =>
            d.title.toLowerCase().includes(q) ||
            d.author.toLowerCase().includes(q)
        );

        if (centuryFilterVal) {
            filtered = filtered.filter(d => d.century === centuryFilterVal);
        }

        const [field, order] = sortKey.split("-");

        filtered.sort((a, b) => {
            let cmp;

            if (field === "century") {
                cmp = (parseInt(a.century) || 0) - (parseInt(b.century) || 0);
            } else {
                cmp = (a[field] || "").localeCompare(b[field] || "");
            }

            return order === "desc" ? -cmp : cmp;
        });

        render(filtered);
    }

    // ── Events ──
    searchInput.addEventListener("input", updateResults);
    sortSelect.addEventListener("change", updateResults);
    filterCentury.addEventListener("change", updateResults);

    clearBtn.addEventListener("click", () => {
        searchInput.value = "";
        sortSelect.value = "title-asc";
        filterCentury.value = "";
        updateResults();
    });

    // ── Init UI ──
    showListView();
    populateCenturyFilter();
    updateResults();

})();