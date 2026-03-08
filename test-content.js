// SharePoint TOC Web Part Test Content Injector
// Use in browser console to test TOC detection with proper headings
// Optimized for one-third left layout

const testContent = document.createElement('div');
testContent.id = 'toc-test-content';
testContent.style.cssText = `
  padding: 40px;
  max-width: 900px;
  margin: 0;
  line-height: 1.6;
  color: #333;
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
`;

testContent.innerHTML = `
  <h2>1. Purpose</h2>
  <p>This section explains the purpose of the DR/BC System Runbook. It provides an overview of why this documentation exists and what it aims to achieve. The runbook serves as a critical reference guide during disaster recovery scenarios.</p>
  
  <h2>2. System Overview</h2>
  <p>This section provides details about the system, ownership, hosting, and importance to the business. Understanding the system architecture is essential for effective disaster recovery planning.</p>
  
  <h3>2.1 Components and Modules</h3>
  <p>List of primary components that make up the system. Each component plays a vital role in the overall system functionality and must be understood for recovery procedures.</p>
  
  <h3>2.2 Hosting &amp; Architecture</h3>
  <p>Information about cloud/on-prem setup, regions, and deployment models. Our system is hosted on Azure with multi-region redundancy for high availability.</p>
  
  <h2>3. Dependencies</h2>
  <p>All internal and external systems this platform interacts with. Understanding these dependencies is critical for identifying potential points of failure.</p>
  
  <h3>3.1 Internal Dependencies</h3>
  <p>Systems within our organization that this depends on. These include identity services, logging infrastructure, and monitoring tools.</p>
  
  <h3>3.2 External Dependencies</h3>
  <p>Third-party services and APIs we rely on. Downtime in these services can impact our system availability.</p>
  
  <h4>3.2.1 Payment Gateway</h4>
  <p>Integration with payment processing services provided by our payment processor. Backup payment methods are available for continuity.</p>
  
  <h4>3.2.2 Authentication Service</h4>
  <p>SSO and authentication provider details. Our system relies on Azure AD for user authentication and authorization.</p>
  
  <h2>4. Business Impact Summary</h2>
  <p>Analysis of how system downtime affects business operations. An hour of downtime could result in significant revenue loss and customer impact.</p>
  
  <h2>5. Current Resiliency State</h2>
  <p>Overview of current disaster recovery capabilities. Our system has achieved a 99.95% uptime SLA with automated failover mechanisms.</p>
  
  <h3>5.1 Backup Status</h3>
  <p>Current backup configuration and schedule. Daily full backups are taken with hourly incremental backups to minimize data loss.</p>
  
  <h3>5.2 Recovery Capabilities</h3>
  <p>What can be recovered and how quickly. RTO is 15 minutes and RPO is 1 hour for critical data.</p>
  
  <h2>6. Recovery Process</h2>
  <p>Step-by-step instructions for disaster recovery. This is the most critical section of the runbook and defines all recovery procedures.</p>
  
  <h3>6.1 Initial Assessment</h3>
  <p>How to assess the scope of the incident. Start by checking monitoring dashboards and alert messages to understand what has failed.</p>
  
  <h3>6.2 Recovery Steps</h3>
  <p>Detailed recovery procedures. Follow these steps in order to restore service as quickly as possible.</p>
  
  <h4>6.2.1 Database Restore</h4>
  <p>Instructions for restoring database from backup. Connect to the backup vault, select the appropriate backup point, and initiate restore to standby instance.</p>
  
  <h4>6.2.2 Application Recovery</h4>
  <p>Steps to bring the application back online. Verify database connectivity, run migration scripts, and perform smoke tests before declaring recovery complete.</p>
  
  <h2>7. Testing &amp; Validation</h2>
  <p>Post-recovery validation procedures to ensure system integrity and data consistency.</p>
  
  <h2>8. Communication Plan</h2>
  <p>Procedures for notifying stakeholders and customers during a disaster event.</p>
`;

// Find the appropriate container in the section (right side of one-third left layout)
// Try multiple selectors to work with different SharePoint versions
const canvas = document.querySelector('[data-automation-id="CanvasZone"]') || 
               document.querySelector('.CanvasZone:not(.tocContainer)') || 
               document.querySelector('[role="main"]') ||
               document.querySelector('main') ||
               document.body;

// Inject content
if (canvas) {
  canvas.appendChild(testContent);
  console.log('✅ Test headings injected successfully!');
  console.log('📋 Your TOC should now show:');
  console.log('  - 8 H2 headings (main sections)');
  console.log('  - 9 H3 headings (subsections)');
  console.log('  - 2 H4 headings (nested details)');
  console.log('\n🔍 Scroll through the content and watch the TOC highlight active sections!');
} else {
  console.error('❌ Could not find canvas element. Make sure you have the TOC web part loaded.');
}

// Optional: Log detected headings for debugging
setTimeout(() => {
  const detectedHeadings = document.querySelectorAll('h2, h3, h4');
  console.log(`\n📊 Total headings on page: ${detectedHeadings.length}`);
  detectedHeadings.forEach((h, i) => {
    console.log(`  ${i + 1}. ${h.tagName} - "${h.textContent.substring(0, 50)}..."`);
  });
}, 500);
