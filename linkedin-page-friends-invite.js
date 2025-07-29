/**
 * Invite friends to follow your page without clicking 250 times (LinkedIn limit. Do not break LinkedIn's terms of use!)
 *
 * This script performs the following:
 * 1. Repeatedly clicks the "Show more results" button until no more profiles load.
 * 2. Selects all visible checkboxes for pending connections.
 * 3. Leaves the "Invite" button for you to click manually (does NOT send automatically).
 *
 * To use:
 * - Open the "Invite connections" dialog on LinkedIn (from your Page).
 * - Open your browser's Developer Console.
 * - Paste and run this script.
 *
 * Note:
 * - Use responsibly. Sending invites too quickly may violate LinkedIn's terms. Their limit is 250 - respect that.
 * - This script avoids scrolling and instead works via button-clicking logic.
 */

(async function inviteConnectionsClickToExpand() {
  console.log("Starting LinkedIn invite script (clicks 'Show more results')");

  const MAX_CLICKS = 30;
  const DELAY = 1000;

  // Helper to wait
  const wait = (ms) => new Promise(res => setTimeout(res, ms));

  // Step 1: Click "Show more results" until it's gone or limit reached
  for (let i = 0; i < MAX_CLICKS; i++) {
    const showMoreBtn = Array.from(document.querySelectorAll('button'))
      .find(btn => btn.innerText.trim().toLowerCase() === 'show more results');
    
    if (!showMoreBtn) {
      console.log("No more 'Show more results' buttons found.");
      break;
    }

    showMoreBtn.click();
    console.log("Clicked 'Show more results' (" + (i + 1) + ")");
    await wait(DELAY);
  }

  // Step 2: Select all checkboxes
  const checkboxes = document.querySelectorAll('input[type="checkbox"]:not(:checked)');
  checkboxes.forEach(checkbox => checkbox.click());
  console.log("Selected " + checkboxes.length + " people.");

  // Step 3: Manual confirmation
  console.log("Please review and manually click 'Invite' to send invitations.");
})();
