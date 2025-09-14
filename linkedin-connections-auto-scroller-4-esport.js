/**
 * LinkedIn Connections Auto-Scroller + CSV Export
 * -----------------------------------
 * Purpose: Export your contacts from LinkedIn.
 * 
 * Important: This follows the speed at which I do this
 *                  manually, saving me clicks. It then rate-
 *.                 limits anyway, to be respectful of LinkedIn.
 *
 * DO NOT GO BEYOND MANUAL SPEED! * DO NOT ABUSE AS AUTOMATION!
 * DO NOT BREAK LINKEDIN TERMS OF SERVICE!
 * 
 * Functionality:
 * - Scrolls & clicks "Load more" on search results for:
 *    LinkedIn/My Network/Connections page path.
 * - Collects profile name, description, date seen, and URL.
 * - Saves one CSV file when stopped (manual or auto).
 * - Pauses 5 min after every 50 clicks to not hit
 *    manual clicking limits, which pause activity.
 *
 * Usage: Paste into Chrome DevTools. 
 * Stop anytime with: window.stopLinkedInScroll()
 *
 * Author: Gadi Evron (with ChatGPT)
 * License: MIT
 * Last Updated: 2025-09-15
 * Version: 0.3
 */

(function() {
  let scrolling = true;
  let clickCount = 0;
  const maxClicks = 50;
  let lastProfileCount = 0;

  const collected = new Map();
  let stallCount = 0;
  const maxStalls = 5;

  function styledLog(msg) {
    console.log(`%c${msg}`, "background: #00ff00; color: black; font-weight: bold; padding: 2px 6px; border-radius: 3px;");
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

  function extractProfiles() {
    const anchors = document.querySelectorAll("a[href*='/in/']");
    anchors.forEach(a => {
      const url = (a.href || "").split("?")[0];
      if (!url) return;
      if (!collected.has(url)) {
        const card = a.closest(".entity-result, .reusable-search__result-container, li") || a;
        const name = card.querySelector("span[aria-hidden='true']")?.innerText?.trim() || "";
        const description = card.querySelector(".entity-result__primary-subtitle, .entity-result__summary, .subline-level-1")?.innerText?.trim() || "";
        const date = new Date().toISOString();
        collected.set(url, { name, description, date, url });
      }
    });
  }

  function saveCSV() {
    if (collected.size === 0) return;
    const esc = s => String(s).replace(/"/g, '""').replace(/\r?\n|\r/g, " ");
    const header = "Name,Description,Date,URL\n";
    const rows = Array.from(collected.values()).map(p =>
      `"${esc(p.name)}","${esc(p.description)}","${esc(p.date)}","${esc(p.url)}"`
    );
    const blob = new Blob([header + rows.join("\n")], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "linkedin_profiles.csv";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  function step() {
    if (!scrolling) return;

    nudgeScroll();
    styledLog(" Nudge scroll step");

    extractProfiles();

    const btn = findLoadMore();
    const profileCount = countProfiles();

    if (btn && clickCount < maxClicks) {
      clickCount++;
      styledLog(` Clicking 'Load more' (${clickCount}/${maxClicks})`);
      btn.click();
      stallCount = 0;
      setTimeout(step, 3000);
      return;
    }

    // --- NEW: if we hit maxClicks, pause 5 minutes then reset ---
    if (btn && clickCount >= maxClicks) {
      styledLog(" Hit 50 clicks — pausing for 5 minutes...");
      clickCount = 0; // reset counter
      setTimeout(step, 5 * 60 * 1000); // sleep 5 minutes
      return;
    }

    if (profileCount === lastProfileCount) {
      styledLog(" No new profiles loaded, nudging harder...");
      window.scrollBy(0, -200);
      stallCount++;
      styledLog(` Stall ${stallCount}/${maxStalls}${btn ? "" : " (no 'Load more' visible)"}`);
      if (stallCount >= maxStalls && (!btn || clickCount >= maxClicks)) {
        styledLog(" Likely end of results — auto-stopping & saving CSV.");
        window.stopLinkedInScroll();
        return;
      }
    } else {
      stallCount = 0;
    }

    lastProfileCount = profileCount;
    setTimeout(step, 2000);
  }

  step();

  window.stopLinkedInScroll = () => {
    scrolling = false;
    styledLog(" Auto-scroll stopped by user.");
    extractProfiles();
    saveCSV();
  };

  styledLog(" Auto-scroll + anti-stall nudges started.");
})();
