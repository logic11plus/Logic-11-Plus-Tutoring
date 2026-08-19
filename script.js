/**
 * Logic 11+ Booking Management & Interactive Script
 * Features:
 * - Form validation and submission
 * - Instant confirmation modal with spam reminder
 * - Secure Google Sheet Web App integration
 */

document.addEventListener('DOMContentLoaded', () => {
  // --- Constants & State ---
  const OFFICIAL_EMAIL = 'logic11plus@gmail.com';

  // Paste your Google Apps Script Web App URL here after deploying (see google-sheets-automation.js)
  const GOOGLE_SHEET_WEBAPP_URL = 'https://script.google.com/macros/s/AKfycbxIydcWoZmumq6C9jQtPpVCyBJ6DjknvLwisH4bePaH7l4_S98X8WNIgEo8THP4ibq2/exec '; // e.g. 'https://script.google.com/macros/s/AKfycb.../exec'

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
  const confirmSlot = document.getElementById('confirmSlot');
  const slotPricingBadge = document.getElementById('slotPricingBadge');

  // Dynamic slot description badge update when changing dropdown
  if (targetYearSelect && slotPricingBadge) {
    targetYearSelect.addEventListener('change', () => {
      const val = targetYearSelect.value;
      if (val.includes('Saturday 9:00 AM')) {
        slotPricingBadge.innerHTML = '<span>Slot: <strong>Saturday 9:00 AM – 10:00 AM</strong></span> • <span>Duration: <strong>1 Hour</strong></span> • <span>Fee: <strong>£15</strong></span>';
      } else if (val.includes('Thursday 6:00 PM')) {
        slotPricingBadge.innerHTML = '<span>Slot: <strong>Thursday 6:00 PM – 7:00 PM</strong></span> • <span>Duration: <strong>1 Hour</strong></span> • <span>Fee: <strong>£15</strong></span>';
      } else if (val.includes('Saturday 10:00 AM')) {
        slotPricingBadge.innerHTML = '<span>Slot: <strong>Saturday 10:00 AM – 11:00 AM</strong></span> • <span>Duration: <strong>1 Hour</strong></span> • <span>Fee: <strong>£15</strong></span>';
      } else if (val.includes('Thursday 5:00 PM')) {
        slotPricingBadge.innerHTML = '<span>Slot: <strong>Thursday 5:00 PM – 6:00 PM</strong></span> • <span>Duration: <strong>1 Hour</strong></span> • <span>Fee: <strong>£15</strong></span>';
      } else {
        slotPricingBadge.innerHTML = '<span>Duration: <strong>1 Hour</strong></span> • <span>Fee: <strong>£15</strong></span>';
      }
    });
  }

  // Mobile menu toggle
  const mobileMenuBtn = document.getElementById('mobileMenuBtn');
  const navLinks = document.getElementById('navLinks');

  if (mobileMenuBtn && navLinks) {
    mobileMenuBtn.addEventListener('click', () => {
      navLinks.classList.toggle('mobile-open');
    });
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
      termsCheckError.textContent = "Please confirm the booking registration checkbox.";
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
        timestamp: new Date().toISOString(),
        formattedDate: new Date().toLocaleString(),
        parentName: parentNameInput.value.trim(),
        childName: childNameInput.value.trim(),
        childSchool: childSchoolInput.value.trim(),
        parentEmail: parentEmailInput.value.trim(),
        targetYear: targetYearSelect.value,
        notes: additionalNotesInput.value.trim() || 'None'
      };

      // 1. Transmit to Google Sheet (if WebApp URL is configured)
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

      // 2. Populate confirmation modal
      if (confirmParentName) confirmParentName.textContent = newLead.parentName;
      if (confirmChildName) confirmChildName.textContent = newLead.childName;
      if (confirmChildSchool) confirmChildSchool.textContent = newLead.childSchool;
      if (confirmEmail) confirmEmail.textContent = newLead.parentEmail;
      if (confirmSlot) confirmSlot.textContent = newLead.targetYear;

      // 3. Show confirmation modal
      if (confirmationModal) {
        confirmationModal.style.display = 'flex';
        setTimeout(() => confirmationModal.classList.add('active'), 10);
      }

      // 4. Reset form fields
      bookingForm.reset();
    });
  }

  // --- Close Confirmation Modal (requires spam checkbox ticked first) ---
  const spamAckCheck = document.getElementById('spamAckCheck');
  const spamAckLabel = document.getElementById('spamAckLabel');

  // Enable the close button only when checkbox is ticked
  if (spamAckCheck && closeConfirmModalBtn) {
    spamAckCheck.addEventListener('change', () => {
      if (spamAckCheck.checked) {
        closeConfirmModalBtn.disabled = false;
        closeConfirmModalBtn.classList.remove('btn-disabled');
        if (spamAckLabel) spamAckLabel.classList.add('ack-checked');
      } else {
        closeConfirmModalBtn.disabled = true;
        closeConfirmModalBtn.classList.add('btn-disabled');
        if (spamAckLabel) spamAckLabel.classList.remove('ack-checked');
      }
    });
  }

  if (closeConfirmModalBtn && confirmationModal) {
    closeConfirmModalBtn.addEventListener('click', () => {
      if (closeConfirmModalBtn.disabled) return;
      confirmationModal.classList.remove('active');
      setTimeout(() => {
        confirmationModal.style.display = 'none';
        // Reset checkbox for next submission
        if (spamAckCheck) spamAckCheck.checked = false;
        if (closeConfirmModalBtn) closeConfirmModalBtn.disabled = true;
        if (spamAckLabel) spamAckLabel.classList.remove('ack-checked');
      }, 200);
    });
  }

  // Close confirmation modal on background click (only if checkbox was already ticked)
  if (confirmationModal) {
    confirmationModal.addEventListener('click', (e) => {
      if (e.target === confirmationModal && !closeConfirmModalBtn.disabled) {
        confirmationModal.classList.remove('active');
        setTimeout(() => { confirmationModal.style.display = 'none'; }, 200);
      }
    });
  }

  // --- Email Enquiry Copy Popup ---
  const emailEnquiryBtn = document.getElementById('emailEnquiryBtn');
  const emailEnquiryModal = document.getElementById('emailEnquiryModal');
  const closeEmailEnquiryBtn = document.getElementById('closeEmailEnquiryBtn');
  const copyEmailBtn = document.getElementById('copyEmailBtn');
  const copyBtnText = document.getElementById('copyBtnText');

  function openEmailModal() {
    if (emailEnquiryModal) {
      emailEnquiryModal.style.display = 'flex';
      setTimeout(() => emailEnquiryModal.classList.add('active'), 10);
    }
  }

  function closeEmailModal() {
    if (emailEnquiryModal) {
      emailEnquiryModal.classList.remove('active');
      setTimeout(() => { emailEnquiryModal.style.display = 'none'; }, 200);
      // Reset copy button text
      if (copyBtnText) copyBtnText.textContent = 'Copy';
    }
  }

  if (emailEnquiryBtn) {
    emailEnquiryBtn.addEventListener('click', openEmailModal);
  }

  if (closeEmailEnquiryBtn) {
    closeEmailEnquiryBtn.addEventListener('click', closeEmailModal);
  }

  // Close on backdrop click
  if (emailEnquiryModal) {
    emailEnquiryModal.addEventListener('click', (e) => {
      if (e.target === emailEnquiryModal) closeEmailModal();
    });
  }

  // Copy email to clipboard
  if (copyEmailBtn) {
    copyEmailBtn.addEventListener('click', () => {
      const email = 'logic11plus@gmail.com';
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(email).then(() => {
          copyBtnText.textContent = 'Copied!';
          copyEmailBtn.classList.add('btn-copy--success');
          setTimeout(() => {
            copyBtnText.textContent = 'Copy';
            copyEmailBtn.classList.remove('btn-copy--success');
          }, 2000);
        });
      } else {
        // Fallback for older browsers
        const tmp = document.createElement('textarea');
        tmp.value = email;
        tmp.style.position = 'fixed';
        tmp.style.opacity = '0';
        document.body.appendChild(tmp);
        tmp.focus();
        tmp.select();
        document.execCommand('copy');
        document.body.removeChild(tmp);
        copyBtnText.textContent = 'Copied!';
        setTimeout(() => { copyBtnText.textContent = 'Copy'; }, 2000);
      }
    });
  }
});

