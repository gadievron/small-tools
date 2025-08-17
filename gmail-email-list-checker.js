/**
 * Gmail Suspicious Email List Checker
 * Simple Apps Script to check if you've interacted over Gmail with the listed suspicious email addresses
 * 
 * Instructions: Paste your email list at the bottom where it says "PASTE EMAILS HERE"
 */

function checkSuspiciousEmails() {
  // Get emails from the list at bottom of script
  const emailText = getSuspiciousEmailText();
  const suspiciousEmails = emailText
    .split('\n')
    .map(email => email.trim())
    .filter(email => email.length > 0 && email.includes('@'));

  console.log(`🔍 Checking ${suspiciousEmails.length} suspicious emails...`);
  
  const results = [];
  let foundCount = 0;
  
  for (let i = 0; i < suspiciousEmails.length; i++) {
    const email = suspiciousEmails[i].toLowerCase();
    
    if (i % 100 === 0) {
      console.log(`Progress: ${i}/${suspiciousEmails.length} (${foundCount} found)`);
    }
    
    try {
      const threads = GmailApp.search(`from:${email} OR to:${email}`, 0, 10);
      const hasInteraction = threads.length > 0;
      
      if (hasInteraction) {
        foundCount++;
        const lastDate = threads.length > 0 ? threads[0].getLastMessageDate() : null;
        console.log(`🚨 FOUND: ${email} (${threads.length} threads)`);
        
        results.push({
          email: email,
          found: true,
          threadCount: threads.length,
          lastDate: lastDate
        });
      }
      
      Utilities.sleep(500);
      
    } catch (error) {
      console.log(`Error checking ${email}: ${error.message}`);
    }
  }
  
  // Generate report
  console.log('\n' + '='.repeat(50));
  console.log('🔍 SUSPICIOUS EMAIL REPORT');
  console.log('='.repeat(50));
  console.log(`Total checked: ${suspiciousEmails.length}`);
  console.log(`Interactions found: ${foundCount}`);
  
  if (foundCount === 0) {
    console.log('\n✅ Good news! No suspicious email interactions found.');
  } else {
    console.log('\n⚠️ WARNING: Found interactions with suspicious emails:');
    results.forEach((result, index) => {
      console.log(`${index + 1}. ${result.email} (${result.threadCount} threads)`);
      if (result.lastDate) {
        console.log(`   Last contact: ${result.lastDate.toDateString()}`);
      }
    });
    console.log('\nRecommendation: Review these interactions for potential security risks.');
  }
  
  return results;
}

// Simple test function for individual emails
function testEmail(email) {
  try {
    const threads = GmailApp.search(`from:${email} OR to:${email}`, 0, 5);
    console.log(`${email}: ${threads.length > 0 ? 'FOUND' : 'Not found'} (${threads.length} threads)`);
    return threads.length > 0;
  } catch (error) {
    console.log(`Error: ${error.message}`);
    return false;
  }
}

// This function contains the email list - paste your emails below
function getSuspiciousEmailText() {
  return `
PASTE EMAILS HERE
ONE EMAIL PER LINE
NO QUOTES OR COMMAS NEEDED
example1@gmail.com
example2@domain.com
`;
}
