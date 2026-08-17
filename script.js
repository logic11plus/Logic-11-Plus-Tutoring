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
  const GOOGLE_SHEET_WEBAPP_URL = ''; // e.g. 'https://script.google.com/macros/s/AKfycb.../exec'

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

      // 3. Show confirmation modal
      if (confirmationModal) {
        confirmationModal.style.display = 'flex';
        setTimeout(() => confirmationModal.classList.add('active'), 10);
      }

      // 4. Reset form fields
      bookingForm.reset();
    });
  }

  // --- Close Confirmation Modal ---
  if (closeConfirmModalBtn && confirmationModal) {
    closeConfirmModalBtn.addEventListener('click', () => {
      confirmationModal.classList.remove('active');
      setTimeout(() => { confirmationModal.style.display = 'none'; }, 200);
    });
  }

  // Close modal on background click
  if (confirmationModal) {
    confirmationModal.addEventListener('click', (e) => {
      if (e.target === confirmationModal) {
        confirmationModal.classList.remove('active');
        setTimeout(() => { confirmationModal.style.display = 'none'; }, 200);
      }
    });
  }
});
