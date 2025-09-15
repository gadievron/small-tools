/**
 * LinkedIn Connections Auto-Scroller + CSV Export
 * -----------------------------------
 * Purpose: Export your contacts from LinkedIn.
 * 
 * Important: This follows the speed at which I do this
 *                  manually, saving me clicks. It then rate-
 *.                 limits anyway, to be respectful of LinkedIn.
 *
 * DO NOT GO BEYOND MANUAL SPEED!
 * DO NOT BREAK LINKEDIN TERMS OF SERVICE!
 * 
 * Functionality:
 * - Scrolls & clicks "Load more" on search results for:
 *    LinkedIn/My Network/Connections page path.
 * - Collects profile name, description, date, and URL.
 * - Saves a test CSV file at the offset, then another CSV
 *    file when stopped (manual or auto).
 *
 * Usage: Paste into Chrome DevTools. 
 * Stop anytime with: window.stopLinkedInScroll()
 *
 * Author: Gadi Evron (with ChatGPT)
 * License: MIT
 * Last Updated: 2025-09-15
 * Version: 0.6
 */

(function() {
  let scrolling = true;
  let clickCount = 0;
  const maxClicks = 50;
  let lastProfileCount = 0;

  const collected = new Map();
  let stallCount = 0;
  const maxStalls = 5;

  // 5–6s jitter between steps
  const nextDelay = () => 5000 + Math.floor(Math.random() * 1001);

  // progressive end probes: 1, 2, 5, 10, 10 minutes
  const stallBackoffMinutes = [1, 2, 5, 10, 10];
  let endProbeIndex = 0;

  // one-time test save after first "Load more"
  let didTestSave = false;

  function styledLog(msg) {
    console.log(
      `%c${msg}`,
      "background: #00ff00; color: black; font-weight: bold; padding: 2px 6px; border-radius: 3px;"
    );
  }

  function countProfiles() {
    return document.querySelectorAll("a[href*='/in/']").length;
  }

  function findLoadMore() {
    return Array.from(document.querySelectorAll("button"))
      .find(btn => btn.innerText.trim().toLowerCase() === "load more");
  }

  function nudgeScroll() {
    window.scrollBy(0, 400);
    window.dispatchEvent(new WheelEvent("wheel", { deltaY: 400, bubbles: true, cancelable: true }));
  }

  // === ONLY data extraction changed earlier; now surgically expand DATE lookup ===
  function extractProfiles() {
    const ACTION_TEXT = /^(message|follow|connect|remove|pending|inmail|open profile|message anyone)$/i;

    // Match your Connections DOM: <p><a href="/in/...">Name</a></p> then <p>Headline</p>
    const nameLinks = document.querySelectorAll("p > a[href*='/in/']");
    nameLinks.forEach(link => {
      const nameText = (link.textContent || "").trim();
      if (!nameText || ACTION_TEXT.test(nameText.toLowerCase())) return;

      const url = (link.href || "").split("?")[0];
      if (!url || collected.has(url)) return;

      const nameP = link.closest("p");
      const innerCard =
        link.closest("[data-view-name='connections-profile']") ||
        nameP?.parentElement || link.parentElement;

      // Description is the very next <p> sibling after the name <p>
      let description = "";
      if (nameP && nameP.nextElementSibling && nameP.nextElementSibling.tagName === "P") {
        description = (nameP.nextElementSibling.textContent || "").trim();
      }
      // fallback inside card (rare)
      if (!description && innerCard) {
        const maybe = innerCard.querySelector("p + p");
        if (maybe) description = (maybe.textContent || "").trim();
      }

      // --- DATE: broaden search to the outer row wrapper (does NOT change other logic) ---
      // Many connections rows are wrapped in a container like div[componentkey^="auto-component-..."].
      const outerRow =
        link.closest("div[componentkey^='auto-component-']") ||
        innerCard ||
        nameP?.parentElement ||
        link.parentElement;

      // Look for visible texts like "Connected on …", "Connected …", "Added on …", "Added …"
      const DATE_RX = /^(connected on|connected|added on|added)\b/i;
      let date = "";

      // search in a few plausible containers: inner card, then outer row, then their parents
      const scan = [];
      if (innerCard) scan.push(innerCard);
      if (outerRow && outerRow !== innerCard) scan.push(outerRow);
      if (outerRow?.parentElement) scan.push(outerRow.parentElement);
      if (innerCard?.parentElement) scan.push(innerCard.parentElement);

      for (const scope of scan) {
        if (!scope) continue;
        // prioritize <time>, then any small text nodes
        const timeEl = scope.querySelector("time");
        if (timeEl) {
          const t = (timeEl.textContent || "").trim();
          if (DATE_RX.test(t)) { date = t; break; }
        }
        const hit = Array.from(scope.querySelectorAll("span, div, time, p"))
          .map(el => (el.textContent || "").trim())
          .find(t => DATE_RX.test(t));
        if (hit) { date = hit; break; }
      }
      // If still empty, leave blank (we DO NOT invent extraction time)

      // Skip obvious garbage rows
      if (!nameText && !description) return;

      collected.set(url, { name: nameText, description, date, url });
    });

    // Safety net for other layouts (search results, etc.) — unchanged
    if (collected.size === 0) {
      document.querySelectorAll("a[href*='/in/']").forEach(a => {
        const url = (a.href || "").split("?")[0];
        if (!url || collected.has(url)) return;

        const raw = (a.textContent || "").trim();
        const isButtonish = a.getAttribute("role") === "button" || !!a.closest("button");
        if (isButtonish || (raw && ACTION_TEXT.test(raw.toLowerCase()))) return;

        let name = raw.split("\n")[0] || "";
        let description = "";

        const p = a.closest("p");
        if (p && p.nextElementSibling && p.nextElementSibling.tagName === "P") {
          description = (p.nextElementSibling.textContent || "").trim();
        }

        // do NOT invent a date here
        collected.set(url, { name, description, date: "", url });
      });
    }
  }

  function saveCSV(forceEvenIfEmpty = false, filename = "linkedin_profiles.csv") {
    if (!forceEvenIfEmpty && collected.size === 0) return;
    const esc = s => String(s ?? "").replace(/"/g, '""').replace(/\r?\n|\r/g, " ");
    const header = "Name,Description,Date,URL\n";
    const rows = Array.from(collected.values()).map(p =>
      `"${esc(p.name)}","${esc(p.description)}","${esc(p.date)}","${esc(p.url)}"`
    );
    const blob = new Blob([header + rows.join("\n")], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  function step() {
    if (!scrolling) return;

    nudgeScroll();
    styledLog("↕️ Nudge scroll step");

    extractProfiles();

    const btn = findLoadMore();
    const profileCount = countProfiles();

    if (btn && clickCount < maxClicks) {
      clickCount++;
      styledLog(`👉 Clicking 'Load more' (${clickCount}/${maxClicks})`);
      btn.click();
      stallCount = 0;

      // one-time test CSV ~5s after first click
      if (!didTestSave && clickCount === 1) {
        styledLog("💾 Test CSV save after first 'Load more' click.");
        setTimeout(() => {
          extractProfiles();
          saveCSV(true, "linkedin_profiles_test.csv");
        }, 5000);
        didTestSave = true;
      }

      setTimeout(step, nextDelay());
      return;
    }

    // pause after 50 clicks — 3 minutes
    if (btn && clickCount >= maxClicks) {
      styledLog("😴 Hit 50 clicks — pausing for 3 minutes...");
      clickCount = 0; // reset counter
      setTimeout(step, 3 * 60 * 1000);
      return;
    }

    if (profileCount === lastProfileCount) {
      styledLog("⚠️ No new profiles loaded, nudging harder...");
      window.scrollBy(0, -200);
      stallCount++;
      styledLog(`⏳ Stall ${stallCount}/${maxStalls}${btn ? "" : " (no 'Load more' visible)"}`);

      // progressive backoff probes before final stop
      if (stallCount >= maxStalls && (!btn || clickCount >= maxClicks)) {
        if (endProbeIndex < stallBackoffMinutes.length) {
          const mins = stallBackoffMinutes[endProbeIndex];
          styledLog(`🕒 Possible end — waiting ${mins} min before retry #${endProbeIndex + 1}/${stallBackoffMinutes.length}...`);
          endProbeIndex++;
          setTimeout(step, mins * 60 * 1000);
          return;
        }
        styledLog("✅ Confirmed end after retries — auto-stopping & saving CSV.");
        window.stopLinkedInScroll();
        return;
      }
    } else {
      stallCount = 0;
      endProbeIndex = 0; // reset probes on progress
    }

    lastProfileCount = profileCount;
    setTimeout(step, nextDelay());
  }

  // run
  step();

  window.stopLinkedInScroll = () => {
    scrolling = false;
    styledLog("⏹️ Auto-scroll stopped by user.");
    extractProfiles();
    saveCSV();
  };

  styledLog("▶️ Auto-scroll + anti-stall nudges started.");
})();
