// LinkedIn Event Smart Invitation Script
//
// This script invites connections to LinkedIn events, and filters by specific keywords
// so you only invite relevant people instead of spamming. It acts like a user would,
// pressing "more" and checking boxes.
//
// IMPORTANT: Do not break LinkedIn's terms of use. Use at your own risk.
// IMPORTANT 2: LinkedIn only allows about 1000 invites to be scrolled and you will
// invite less when using this script.
//
// HOW TO CUSTOMIZE THIS SCRIPT:
//
// All filtering in this script is done by **substring matching** on the LinkedIn title line.
// Any part of a job title or company name matching your strings will trigger a match.
//
// Matching is case-insensitive.
//
// EXAMPLES OF HOW SUBSTRING MATCHING WORKS:
//   - "Boss"       → matches "Big Boss", "Boss of Nothing", "Boss-level Engineer"
//   - "Crypto Bro" → matches "Senior Crypto Bro", "Ex-Crypto Bro Founder"
//   - "ACME"       → matches "ACME Corp", "ACME Finance"
//   - "VP"         → matches "VP Sales", "VP of Nothing"
//
// HOW TO CONFIGURE:
//
// 1) config.requiredKeywords (ALLOW LIST)
//    - These substrings MUST appear somewhere in the title for the person to be invited.
//    - Default examples: "Boss", "Crypto Bro".
//    - Add/remove as many as you want (one string per line).
//
// 2) config.excludeCompanies (DENY LIST)
//    - If the title/company line contains any of these substrings, the person is skipped.
//    - Default examples: "ACME", "OCP".
//    - Add/remove entries freely.
//
// 3) config.excludeRoles (DENY LIST)
//    - If the title contains any substring from this list, that person is skipped.
//    - Default examples: "VP", "Intern".
//    - Add/remove entries freely.
//
// HOW TO RUN:
// - Open the LinkedIn event invite page (with the invite list visible).
// - Open DevTools → Console, paste this entire script, press Enter.
// 
// By: Gadi Evron (with ChatGPT)
// License: MIT

(async function() {
    'use strict';

    const config = {
        scrollAmount: 400,
        waitAfterScroll: 1000,
        waitBetweenClicks: 80,
        maxScrollAttempts: 100,
        stuckDelays: [30000, 60000, 180000],
        
        // ALLOW LIST (must match at least one substring in the title)
        requiredKeywords: [
            'Boss',        // example: matches any title containing "Boss"
            'Crypto Bro'   // example: matches any title containing "Crypto Bro"
            // Add more allow-list substrings here...
        ],
        
        // DENY LIST (company substrings to exclude)
        excludeCompanies: [
            'ACME',        // example: matches "ACME", "ACME Corp", etc.
            'OCP'          // example: matches "OCP", "OCP Systems", etc.
            // Add more company deny-list substrings here...
        ],
        
        // DENY LIST (role substrings to exclude)
        excludeRoles: [
            'VP',          // example: matches "VP of X", "Senior VP", etc.
            'Intern'       // example: matches "Intern", "Summer Intern", etc.
            // Add more role deny-list substrings here...
        ]
    };

    const stats = {
        checked: 0,
        skipped: 0,
        alreadyInvited: 0,
        scrolls: 0
    };

    const processedCheckboxes = new Set();

    console.log('═════════════════════════════════════════════════');
    console.log('LinkedIn Smart Event Invite Script');
    console.log('═════════════════════════════════════════════════');
    console.log('Required Keywords (substring allow list):');
    config.requiredKeywords.forEach(kw => console.log(`   - "${kw}"`));
    console.log(`\nCompany Deny List: ${config.excludeCompanies.length} entries`);
    console.log('Role Deny List:', config.excludeRoles.join(', '));
    console.log('─────────────────────────────────────────────────\n');

    function wait(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    function getScrollContainer() {
        return document.querySelector('#invitee-picker-results-container');
    }

    function clickShowMoreButton() {
        const showMoreButton = document.querySelector('button.scaffold-finite-scroll__load-button');
        if (showMoreButton && showMoreButton.offsetParent !== null) {
            console.log('📄 Clicking "Show more results" button...');
            showMoreButton.click();
            return true;
        }
        return false;
    }

    async function checkAllUncheckedBoxes() {
        const listItems = document.querySelectorAll('.artdeco-typeahead__result.ember-view');
        
        console.log(`\n📊 Found ${listItems.length} people in the list`);
        
        for (const listItem of listItems) {

            const statusDiv = listItem.querySelector('.invitee-picker-connections-result-item__status');
            if (statusDiv && statusDiv.textContent.trim() === 'Invited') {
                stats.alreadyInvited++;
                continue;
            }

            const checkbox = listItem.querySelector('input[type="checkbox"].ember-checkbox');
            if (!checkbox || checkbox.disabled || processedCheckboxes.has(checkbox)) {
                continue;
            }

            const nameElement = listItem.querySelector('.t-16.t-black.t-bold');
            const name = nameElement ? nameElement.textContent.trim() : 'Unknown';

            const titleElement = listItem.querySelector('.t-14.t-black--light.t-normal');
            const title = titleElement ? titleElement.textContent.trim().toLowerCase() : '';
            
            console.log(`🔍 Checking: ${name}`);
            console.log(`   Title: "${title}"`);
            
            // ALLOW LIST: must match at least one substring
            const matchedKeyword = config.requiredKeywords.find(keyword => 
                title.includes(keyword.toLowerCase())
            );
            
            if (!matchedKeyword) {
                console.log('   ⊗ SKIP: No required substring matched');
                stats.skipped++;
                continue;
            }
            
            console.log(`   ✓ Matched allow substring: "${matchedKeyword}"`);
            
            // COMPANY DENY LIST
            const matchedCompany = config.excludeCompanies.find(company => 
                title.includes(company.toLowerCase())
            );
            if (matchedCompany) {
                console.log(`   ⊗ SKIP: Company deny substring matched "${matchedCompany}"`);
                stats.skipped++;
                continue;
            }
            
            // ROLE DENY LIST
            const matchedRole = config.excludeRoles.find(role => 
                title.includes(role.toLowerCase())
            );
            if (matchedRole) {
                console.log(`   ⊗ SKIP: Role deny substring matched "${matchedRole}"`);
                stats.skipped++;
                continue;
            }

            // INVITE
            if (!checkbox.checked) {
                checkbox.click();
                processedCheckboxes.add(checkbox);
                console.log(`   ✅ INVITED: ${name}`);
                stats.checked++;
                await wait(config.waitBetweenClicks);
            }
        }
    }

    async function autoScroll() {
        const container = getScrollContainer();
        
        if (!container) {
            console.error('❌ Could not find scroll container');
            return;
        }

        let scrollAttempts = 0;
        let lastScrollHeight = container.scrollHeight;
        let stuckCount = 0;
        let noNewContentCount = 0;

        while (scrollAttempts < config.maxScrollAttempts) {
            await checkAllUncheckedBoxes();

            if (clickShowMoreButton()) {
                await wait(config.waitAfterScroll * 2);
            }

            container.scrollTop += config.scrollAmount;
            stats.scrolls++;
            scrollAttempts++;

            console.log(`📜 Scroll ${scrollAttempts}`);

            await wait(config.waitAfterScroll);

            const newScrollHeight = container.scrollHeight;
            
            if (newScrollHeight === lastScrollHeight) {
                noNewContentCount++;
                console.log(`⏸️ No new content (${noNewContentCount})`);

                if (container.scrollTop + container.clientHeight >= container.scrollHeight - 10) {
                    if (stuckCount < config.stuckDelays.length) {
                        const waitTime = config.stuckDelays[stuckCount];
                        console.log(`⏳ Waiting ${waitTime / 1000}s...`);
                        await wait(waitTime);
                        stuckCount++;
                        
                        if (clickShowMoreButton()) {
                            await wait(config.waitAfterScroll * 2);
                            lastScrollHeight = container.scrollHeight;
                            noNewContentCount = 0;
                            continue;
                        }
                    } else {
                        console.log('✅ End of list.');
                        break;
                    }
                }

                if (noNewContentCount >= 5) {
                    console.log('✅ No more content after multiple attempts.');
                    break;
                }

            } else {
                lastScrollHeight = newScrollHeight;
                noNewContentCount = 0;
                stuckCount = 0;
            }
        }

        console.log('\n📋 Final check...');
        await checkAllUncheckedBoxes();

        console.log('\n═════════════════════════════════════════════════');
        console.log('✅ Script complete');
        console.log('─────────────────────────────────────────────────');
        console.log(`✓ Invited: ${stats.checked}`);
        console.log(`⊗ Skipped: ${stats.skipped}`);
        console.log(`ℹ️ Already invited: ${stats.alreadyInvited}`);
        console.log(`↻ Scrolls: ${stats.scrolls}`);
        console.log('═════════════════════════════════════════════════\n');
    }

    autoScroll();
})();
