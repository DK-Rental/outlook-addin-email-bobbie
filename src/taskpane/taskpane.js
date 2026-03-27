/*
 * DK Rental Outlook Add-in
 * Property Management Integration
 */

/* global document, Office */

let currentEmail = {
    subject: '',
    sender: '',
    body: '',
    receivedTime: null
};

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        // Hide sideload message, show app
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
        
        // Load email details
        loadEmailDetails();
        
        // Set up event listeners
        setupEventListeners();
        
        // Set up regenerate button (moved inside Office.onReady)
        const regenerateBtn = document.getElementById("regenerate-btn");
        if (regenerateBtn) {
            regenerateBtn.onclick = () => regenerateContent(false); // ✅ manual click
        }

        // Optional keyboard shortcut
        document.addEventListener('keydown', (e) => {
            if ((e.metaKey || e.ctrlKey) && e.key === 'r') {
                e.preventDefault();
                regenerateContent(false); // ✅ manual click
            }
        });
    }
});

function setupEventListeners() {
    // Action items
    document.getElementById("link-property").onclick = linkToProperty;
    document.getElementById("schedule-viewing").onclick = scheduleViewing;
    document.getElementById("process-payment").onclick = processPayment;
    document.getElementById("tenant-info").onclick = showTenantInfo;
    
    // Buttons
    document.getElementById("save-to-property").onclick = saveEmailToProperty;
}


function checkPropertyLink() {
    // This would check your company's database
    // For demo, we'll show a mock property if email contains certain keywords
    const item = Office.context.mailbox.item;
    
    if (item.subject && item.subject.toLowerCase().includes("123 main st")) {
        document.getElementById("property-card").style.display = "block";
        document.getElementById("property-address").textContent = "123 Main Street";
        document.getElementById("property-details").textContent = "3 bed, 2 bath - Current tenant: John Smith";
    } else {
        document.getElementById("property-card").style.display = "none";
    }
}

function linkToProperty() {
    showStatus("Opening property selector...", "success");
    
    // This would open a dialog to select a property
    Office.context.ui.displayDialogAsync(
        'https://localhost:3000/property-selector.html',
        { height: 60, width: 30 },
        function(result) {
            console.log("Dialog opened: " + JSON.stringify(result));
        }
    );
}

function scheduleViewing() {
    const item = Office.context.mailbox.item;
    
    // Create a calendar item from email
    Office.context.mailbox.displayNewAppointmentForm({
        subject: `Viewing: ${item.subject}`,
        location: "Property location",
        start: new Date(),
        end: new Date(new Date().getTime() + 60*60000) // +1 hour
    });
    
    showStatus("Creating viewing appointment...", "success");
}

function processPayment() {
    showStatus("Opening payment processor...", "success");
    
    // Extract potential payment info from email
    const item = Office.context.mailbox.item;
    
    // This would open payment dialog
    alert(`Process payment for: ${item.subject}\nThis would integrate with your payment system.`);
}

function showTenantInfo() {
    // This would fetch tenant info from your database
    const mockTenant = {
        name: "John Smith",
        phone: "(604) 555-0123",
        email: "john.smith@email.com",
        property: "123 Main St",
        leaseStart: "Jan 1, 2025",
        rent: "$2,500/month"
    };
    
    const message = `Tenant Information:
Name: ${mockTenant.name}
Phone: ${mockTenant.phone}
Email: ${mockTenant.email}
Property: ${mockTenant.property}
Lease Start: ${mockTenant.leaseStart}
Rent: ${mockTenant.rent}`;
    
    alert(message);
    showStatus("Tenant info loaded", "success");
}

function saveEmailToProperty() {
    const item = Office.context.mailbox.item;
    
    // This would save to your company database
    console.log("Saving email to property:", {
        subject: item.subject,
        from: item.from ? item.from.emailAddress : "unknown",
        received: item.dateTimeCreated
    });
    
    showStatus("Email saved to property records!", "success");
    
    // Show property card if not visible
    document.getElementById("property-card").style.display = "block";
    document.getElementById("property-address").textContent = "Current Property";
    document.getElementById("property-details").textContent = "Email linked to property records";
}

function showStatus(message, type) {
    const statusEl = document.getElementById("status-message");
    statusEl.textContent = message;
    statusEl.className = `status-message ${type}`;
    
    // Hide after 3 seconds
    setTimeout(() => {
        statusEl.style.display = "none";
    }, 3000);
}

// Export for module usage
export async function run() {
    // Legacy function - now using the new UI
    loadEmailDetails();
}

async function regenerateContent(isFirstLoad = false) {
    const button = document.getElementById("regenerate-btn");
    const contentDiv = document.getElementById("regeneratable-content");
    const insightsDiv = document.getElementById("insights-content");
    
    // ✅ Only update button if manual click
    if (!isFirstLoad) {
        button.disabled = true;
        button.innerHTML = '<span class="emoji">⏳</span><span>Regenerating...</span>';
    }
    contentDiv.classList.add("loading");
    
    try {
        const newContent = await fetchNewData();
        
        insightsDiv.style.opacity = '0';
        
        setTimeout(() => {
            insightsDiv.innerHTML = newContent;
            insightsDiv.style.opacity = '1';
            
            if (isFirstLoad) {
                showStatus("AI analysis complete!", "success");
            } else {
                showStatus("Content regenerated successfully!", "success");
            }
        }, 300);
        
    } catch (error) {
        showStatus("Error generating AI response", "error");
        console.error(error);
        
    } finally {
        setTimeout(() => {
            // ✅ Only restore button if manual click
            if (!isFirstLoad) {
                button.disabled = false;
                button.innerHTML = '<span class="emoji">🔄</span><span>Regenerate Email Analysis</span>';
            }
            contentDiv.classList.remove("loading");
        }, 400);
    }
}

async function fetchNewData() {
    const item = Office.context.mailbox.item;

    const email_body = await new Promise((resolve, reject) => {
        item.body.getAsync(Office.CoercionType.Text, (bodyResult) => {
            if (bodyResult.status === Office.AsyncResultStatus.Failed) {
                reject(bodyResult.error.message);
            } else {
                resolve(bodyResult.value);
            }
        });
    });

    const response = await fetch('https://localhost:5000/api/analyze', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            subject: item.subject,
            body: email_body
        })
    });

    const aiResponse = await response.json();

    // Return HTML string to be injected into insightsDiv
    return `
        <div class="insight-item">
            <span class="insight-label">Category</span>
            <span class="insight-value">${aiResponse.category}</span>
        </div>
        <div class="insight-item">
            <span class="insight-label">Subcategory</span>
            <span class="insight-value">${aiResponse.subcategory}</span>
        </div>
        <div class="insight-item">
            <span class="insight-label">Summary</span>
            <span class="insight-value">${aiResponse.summary}</span>
        </div>
        <div class="insight-item">
            <span class="insight-label">Intent</span>
            <span class="insight-value">${aiResponse.intent}</span>
        </div>
        <div class="insight-item">
            <span class="insight-label">Suggested Action</span>
            <span class="insight-value">${aiResponse.copilot_action}</span>
        </div>
        <div class="insight-item" style="flex-direction: column; gap: 8px;">
            <span class="insight-label">Draft Reply</span>
            <span class="insight-value" style="white-space: pre-wrap;">${aiResponse.draft_reply}</span>
        </div>
    `;
}

// // Optional: Floating button click handler (if you added it)
// const floatingBtn = document.getElementById("floating-regenerate");
// if (floatingBtn) {
//     floatingBtn.onclick = regenerateContent;
// }

function loadEmailDetails() {
    const item = Office.context.mailbox.item;
    
    currentEmail.subject = item.subject;
    document.getElementById("email-subject").textContent = item.subject || "No subject";
    
    if (item.from) {
        currentEmail.sender = item.from.emailAddress || item.from.displayName || "Unknown sender";
        document.getElementById("email-sender").textContent = `From: ${currentEmail.sender}`;
    }
    
    item.body.getAsync("text", { asyncContext: this }, function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            currentEmail.body = result.value.substring(0, 200);
        }
    });
    
    checkPropertyLink();
    regenerateContent(true); // ✅ auto-run AI analysis on load
}