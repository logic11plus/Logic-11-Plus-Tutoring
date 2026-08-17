/**
 * Logic 11+ Booking Management & Interactive Script
 * Features:
 * - Form validation and submission
 * - Instant confirmation modal with spam reminder
 * - Local storage lead capture
 * - Excel / CSV export functionality
 * - Direct email sender workflow (Yes -> Open preformatted email)
 */

document.addEventListener('DOMContentLoaded', () => {
  // --- Constants & State ---
  const STORAGE_KEY = 'logic11plus_leads';
  const OFFICIAL_EMAIL = 'logic11plus@gmail.com';

  // Paste your Google Apps Script Web App URL here after deploying (see google-sheets-automation.js)
  const GOOGLE_SHEET_WEBAPP_URL = ''; // e.g. 'https://script.google.com/macros/s/AKfycbxIydcWoZmumq6C9jQtPpVCyBJ6DjknvLwisH4bePaH7l4_S98X8WNIgEo8THP4ibq2/exec'

  // DOM Elements
  const bookingForm = document.getElementById('tuitionBookingForm');
  const parentNameInput = document.getElementById('parentName');
  const childNameInput = document.getElementById('childName');
  const childSchoolInput = document.getElementById('childSchool');
  const parentEmailInput = document.getElementById('parentEmail');
  const targetYearSelect = document.getElementById('targetYear');
  const additionalNotesInput = document.getElementById('additionalNotes');
  const termsCheck = document.getElementById('termsCheck');

  // Error fields
  const parentNameError = document.getElementById('parentNameError');
  const childNameError = document.getElementById('childNameError');
  const childSchoolError = document.getElementById('childSchoolError');
  const parentEmailError = document.getElementById('parentEmailError');
  const termsCheckError = document.getElementById('termsCheckError');

  // Confirmation Modal Elements
  const confirmationModal = document.getElementById('confirmationModal');
  const closeConfirmModalBtn = document.getElementById('closeConfirmModalBtn');
  const confirmParentName = document.getElementById('confirmParentName');
  const confirmChildName = document.getElementById('confirmChildName');
  const confirmChildSchool = document.getElementById('confirmChildSchool');
  const confirmEmail = document.getElementById('confirmEmail');

  // Admin & Leads Table Elements
  const leadCountDisplay = document.getElementById('leadCountDisplay');
  const downloadExcelBtn = document.getElementById('downloadExcelBtn');
  const viewLeadsModalBtn = document.getElementById('viewLeadsModalBtn');
  const leadsModal = document.getElementById('leadsModal');
  const closeLeadsModalBtn = document.getElementById('closeLeadsModalBtn');
  const leadsTableBody = document.getElementById('leadsTableBody');
  const exportCsvFromModalBtn = document.getElementById('exportCsvFromModalBtn');
  const clearLeadsBtn = document.getElementById('clearLeadsBtn');

  // Mobile menu toggle
  const mobileMenuBtn = document.getElementById('mobileMenuBtn');
  const navLinks = document.querySelector('.nav-links');

  if (mobileMenuBtn && navLinks) {
    mobileMenuBtn.addEventListener('click', () => {
      navLinks.classList.toggle('mobile-open');
    });
  }

  // --- Lead Storage Helpers ---
  function getLeads() {
    try {
      const stored = localStorage.getItem(STORAGE_KEY);
      return stored ? JSON.parse(stored) : [];
    } catch (e) {
      console.error('Error reading localStorage:', e);
      return [];
    }
  }

  function saveLeads(leads) {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(leads));
      updateLeadsUI();
    } catch (e) {
      console.error('Error saving to localStorage:', e);
    }
  }

  function addLead(lead) {
    const leads = getLeads();
    leads.unshift(lead); // Add newest first
    saveLeads(leads);
  }

  // --- Validation ---
  function validateEmail(email) {
    const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return re.test(String(email).toLowerCase());
  }

  function clearErrors() {
    [parentNameError, childNameError, childSchoolError, parentEmailError, termsCheckError].forEach(el => {
      if (el) el.textContent = '';
    });
    [parentNameInput, childNameInput, childSchoolInput, parentEmailInput].forEach(el => {
      if (el) el.classList.remove('input-invalid');
    });
  }

  function validateForm() {
    clearErrors();
    let isValid = true;

    if (!parentNameInput.value.trim()) {
      parentNameError.textContent = "Please enter parent or guardian's full name.";
      parentNameInput.classList.add('input-invalid');
      isValid = false;
    }

    if (!childNameInput.value.trim()) {
      childNameError.textContent = "Please enter your child's name.";
      childNameInput.classList.add('input-invalid');
      isValid = false;
    }

    if (!childSchoolInput.value.trim()) {
      childSchoolError.textContent = "Please enter your child's current school.";
      childSchoolInput.classList.add('input-invalid');
      isValid = false;
    }

    const emailVal = parentEmailInput.value.trim();
    if (!emailVal) {
      parentEmailError.textContent = "Please provide an email address.";
      parentEmailInput.classList.add('input-invalid');
      isValid = false;
    } else if (!validateEmail(emailVal)) {
      parentEmailError.textContent = "Please enter a valid email address.";
      parentEmailInput.classList.add('input-invalid');
      isValid = false;
    }

    if (!termsCheck.checked) {
      termsCheckError.textContent = "Please acknowledge that classes are online group format with no free taster sessions.";
      isValid = false;
    }

    return isValid;
  }

  // --- Form Submission Handler ---
  if (bookingForm) {
    bookingForm.addEventListener('submit', (e) => {
      e.preventDefault();

      if (!validateForm()) {
        return;
      }

      const newLead = {
        id: 'lead_' + Date.now(),
        timestamp: new Date().toISOString(),
        formattedDate: new Date().toLocaleString(),
        parentName: parentNameInput.value.trim(),
        childName: childNameInput.value.trim(),
        childSchool: childSchoolInput.value.trim(),
        parentEmail: parentEmailInput.value.trim(),
        targetYear: targetYearSelect.value,
        notes: additionalNotesInput.value.trim() || 'None',
        emailSent: false
      };

      // 1. Store in local state/leads storage
      addLead(newLead);

      // 2. Transmit to Google Sheet (if WebApp URL is configured)
      if (GOOGLE_SHEET_WEBAPP_URL && GOOGLE_SHEET_WEBAPP_URL.trim() !== '') {
        fetch(GOOGLE_SHEET_WEBAPP_URL, {
          method: 'POST',
          mode: 'no-cors', // standard for Google Apps Script Web Apps
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(newLead)
        }).catch(err => console.log('Google Sheet sync notice:', err));
      }

      // 3. Populate confirmation modal
      confirmParentName.textContent = newLead.parentName;
      confirmChildName.textContent = newLead.childName;
      confirmChildSchool.textContent = newLead.childSchool;
      confirmEmail.textContent = newLead.parentEmail;

      // 4. Show confirmation modal
      confirmationModal.classList.add('active');
      confirmationModal.setAttribute('aria-hidden', 'false');

      // 5. Reset form fields
      bookingForm.reset();
    });
  }

  // --- Close Confirmation Modal ---
  if (closeConfirmModalBtn) {
    closeConfirmModalBtn.addEventListener('click', () => {
      confirmationModal.classList.remove('active');
      confirmationModal.setAttribute('aria-hidden', 'true');
    });
  }

  // --- Admin Modal Open / Close ---
  if (viewLeadsModalBtn) {
    viewLeadsModalBtn.addEventListener('click', () => {
      renderLeadsTable();
      leadsModal.classList.add('active');
      leadsModal.setAttribute('aria-hidden', 'false');
    });
  }

  if (closeLeadsModalBtn) {
    closeLeadsModalBtn.addEventListener('click', () => {
      leadsModal.classList.remove('active');
      leadsModal.setAttribute('aria-hidden', 'true');
    });
  }

  // Close modals on overlay background click
  [confirmationModal, leadsModal].forEach(modal => {
    modal.addEventListener('click', (e) => {
      if (e.target === modal) {
        modal.classList.remove('active');
        modal.setAttribute('aria-hidden', 'true');
      }
    });
  });

  // --- Render Leads Table & Email Trigger Action ---
  function renderLeadsTable() {
    const leads = getLeads();
    leadsTableBody.innerHTML = '';

    if (leads.length === 0) {
      leadsTableBody.innerHTML = `
        <tr>
          <td colspan="8" style="text-align: center; color: #64748b; padding: 2rem;">
            No bookings recorded yet. Submit a test booking using the form above!
          </td>
        </tr>
      `;
      return;
    }

    leads.forEach(lead => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${lead.formattedDate}</td>
        <td><strong>${escapeHtml(lead.parentName)}</strong></td>
        <td>${escapeHtml(lead.childName)}</td>
        <td>${escapeHtml(lead.childSchool)}</td>
        <td><a href="mailto:${escapeHtml(lead.parentEmail)}">${escapeHtml(lead.parentEmail)}</a></td>
        <td>${escapeHtml(lead.targetYear)}</td>
        <td>
          <button class="btn-trigger-email" data-id="${lead.id}">
            ✉️ Yes, Send Email
          </button>
        </td>
        <td>
          <span class="status-badge ${lead.emailSent ? 'status-sent' : 'status-pending'}">
            ${lead.emailSent ? 'Email Sent' : 'Pending'}
          </span>
        </td>
      `;
      leadsTableBody.appendChild(tr);
    });

    // Attach click triggers for "Send Email" action
    document.querySelectorAll('.btn-trigger-email').forEach(btn => {
      btn.addEventListener('click', function() {
        const leadId = this.getAttribute('data-id');
        triggerSendEmail(leadId);
      });
    });
  }

  // --- Send Email Trigger (Logic: Clicking "Yes" initiates email from logic11plus@gmail.com) ---
  function triggerSendEmail(leadId) {
    const leads = getLeads();
    const lead = leads.find(l => l.id === leadId);
    if (!lead) return;

    // Compose formatted pre-filled email from logic11plus@gmail.com
    const subject = encodeURIComponent(`Logic 11+ Tuition: Enrollment Confirmation for ${lead.childName}`);
    const body = encodeURIComponent(
`Dear ${lead.parentName},

Thank you for booking Logic 11+ Mathematics Tuition for ${lead.childName} (${lead.childSchool}).

We have received your registration for our online group sessions (${lead.targetYear}).

Key Details of Your Group Tuition:
- Format: Online Interactive Live Group Classroom
- Subject Focus: Pure 11+ Mathematics & Logical Problem-Solving
- Direct Tutor Contact: ${OFFICIAL_EMAIL}

We will be sending your cohort schedule, lesson times, and digital classroom login link shortly.

If you have any questions or require anything in the meantime, simply reply to this email (${OFFICIAL_EMAIL}).

Kind regards,
The Logic 11+ Team
${OFFICIAL_EMAIL}`
    );

    // Open default mail client / webmail composer addressed to the parent
    const mailtoUrl = `mailto:${lead.parentEmail}?subject=${subject}&body=${body}`;
    window.location.href = mailtoUrl;

    // Mark status as email sent in storage
    lead.emailSent = true;
    saveLeads(leads);
    renderLeadsTable();
  }

  // --- Export to CSV / Excel ---
  function exportLeadsToCsv() {
    const leads = getLeads();
    if (leads.length === 0) {
      alert('No booking records to export yet!');
      return;
    }

    const headers = ['Submission Date', 'Parent Name', 'Child Name', 'Child School', 'Parent Email', 'Target Year', 'Additional Notes', 'Send Further Email', 'Status'];
    
    const rows = leads.map(l => [
      `"${l.formattedDate}"`,
      `"${escapeCsv(l.parentName)}"`,
      `"${escapeCsv(l.childName)}"`,
      `"${escapeCsv(l.childSchool)}"`,
      `"${escapeCsv(l.parentEmail)}"`,
      `"${escapeCsv(l.targetYear)}"`,
      `"${escapeCsv(l.notes)}"`,
      `"${l.emailSent ? 'Yes' : 'Pending'}"`,
      `"${l.emailSent ? 'Sent' : 'Pending Review'}"`
    ]);

    const csvContent = 'data:text/csv;charset=utf-8,\uFEFF' + [headers.join(','), ...rows.map(e => e.join(','))].join('\n');
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement('a');
    link.setAttribute('href', encodedUri);
    link.setAttribute('download', `Logic11Plus_Bookings_${new Date().toISOString().slice(0,10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }

  if (downloadExcelBtn) {
    downloadExcelBtn.addEventListener('click', exportLeadsToCsv);
  }

  if (exportCsvFromModalBtn) {
    exportCsvFromModalBtn.addEventListener('click', exportLeadsToCsv);
  }

  // --- Clear Leads ---
  if (clearLeadsBtn) {
    clearLeadsBtn.addEventListener('click', () => {
      if (confirm('Are you sure you want to clear all logged booking records?')) {
        localStorage.removeItem(STORAGE_KEY);
        updateLeadsUI();
        renderLeadsTable();
      }
    });
  }

  // --- Update Leads Count UI ---
  function updateLeadsUI() {
    const leads = getLeads();
    if (leadCountDisplay) {
      leadCountDisplay.textContent = `${leads.length} booking${leads.length === 1 ? '' : 's'} logged`;
    }
  }

  // Utility to prevent XSS
  function escapeHtml(str) {
    if (!str) return '';
    return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;");
  }

  function escapeCsv(str) {
    if (!str) return '';
    return str.toString().replace(/"/g, '""');
  }

  // Initial UI refresh
  updateLeadsUI();
});
