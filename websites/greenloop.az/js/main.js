// GreenLoop — shared front-end behaviour (no build step, vanilla JS)

document.addEventListener('DOMContentLoaded', function () {

  /* Mobile nav toggle */
  var burger = document.querySelector('.burger');
  var nav = document.querySelector('.main-nav');
  if (burger && nav) {
    burger.addEventListener('click', function () {
      nav.classList.toggle('nav-open');
    });
  }

  /* Buyer/Seller account switcher: click-to-open (in addition to hover) for touch devices */
  document.querySelectorAll('.account-switch > .icon-btn').forEach(function (icon) {
    icon.addEventListener('click', function (e) {
      e.preventDefault();
      var wrap = icon.closest('.account-switch');
      var alreadyOpen = wrap.classList.contains('open');
      document.querySelectorAll('.account-switch.open').forEach(function (w) { w.classList.remove('open'); });
      if (!alreadyOpen) wrap.classList.add('open');
    });
  });
  document.addEventListener('click', function (e) {
    if (!e.target.closest('.account-switch')) {
      document.querySelectorAll('.account-switch.open').forEach(function (w) { w.classList.remove('open'); });
    }
  });

  /* Generic tab system: [data-tabs] wraps .tab-btn[data-tab] + .tab-panel[data-panel] */
  document.querySelectorAll('[data-tabs]').forEach(function (wrap) {
    var btns = wrap.querySelectorAll('.tab-btn');
    var panels = wrap.querySelectorAll('.tab-panel');
    btns.forEach(function (btn) {
      btn.addEventListener('click', function () {
        btns.forEach(function (b) { b.classList.remove('active'); });
        panels.forEach(function (p) { p.classList.remove('active'); });
        btn.classList.add('active');
        var target = wrap.querySelector('.tab-panel[data-panel="' + btn.dataset.tab + '"]');
        if (target) target.classList.add('active');
      });
    });
  });

  /* Filter chip toggle (visual only) */
  document.querySelectorAll('.chip[data-toggle]').forEach(function (chip) {
    chip.addEventListener('click', function () {
      chip.classList.toggle('active');
    });
  });

  /* Radio tile selection styling handled by :has() in CSS, but keep a JS fallback */
  document.querySelectorAll('.radio-tile input').forEach(function (input) {
    input.addEventListener('change', function () {
      var group = input.closest('.radio-tiles');
      if (!group) return;
      group.querySelectorAll('.radio-tile').forEach(function (t) { t.classList.remove('is-checked'); });
      input.closest('.radio-tile').classList.add('is-checked');
    });
  });

  /* Toggle filter sidebar on mobile */
  var filterToggle = document.querySelector('[data-filter-toggle]');
  var filterPanel = document.querySelector('[data-filter-panel]');
  if (filterToggle && filterPanel) {
    filterToggle.addEventListener('click', function () {
      filterPanel.classList.toggle('is-open');
    });
  }

  /* Prevent demo forms from navigating away; show a lightweight confirmation */
  document.querySelectorAll('form[data-demo-form]').forEach(function (form) {
    form.addEventListener('submit', function (e) {
      e.preventDefault();
      var note = form.querySelector('.form-note');
      if (note) {
        note.textContent = form.dataset.successMessage || 'Submitted — thank you!';
        note.style.display = 'block';
      } else {
        alert(form.dataset.successMessage || 'Submitted — thank you!');
      }
    });
  });

  /* Star-rating input widget: [data-rating-input] */
  document.querySelectorAll('[data-rating-input]').forEach(function (widget) {
    var stars = widget.querySelectorAll('span');
    stars.forEach(function (star, idx) {
      star.addEventListener('click', function () {
        stars.forEach(function (s, i) {
          s.textContent = i <= idx ? '★' : '☆';
        });
        widget.dataset.value = idx + 1;
      });
    });
  });

  /* Simple accordion: [data-accordion] > [data-acc-item] > [data-acc-trigger] + [data-acc-body] */
  document.querySelectorAll('[data-acc-trigger]').forEach(function (trigger) {
    trigger.addEventListener('click', function () {
      var item = trigger.closest('[data-acc-item]');
      item.classList.toggle('open');
    });
  });

  /* =========================================================
     Marketplace filtering (marketplace.html)
     ========================================================= */
  var grid = document.getElementById('listing-grid');
  if (grid) {
    var cards = Array.prototype.slice.call(grid.querySelectorAll('.listing-card'));
    var catBoxes = Array.prototype.slice.call(document.querySelectorAll('[data-filter="category"]'));
    var gradeChips = Array.prototype.slice.call(document.querySelectorAll('.chip[data-grade]'));
    var certBox = document.getElementById('filter-certified');
    var weightMin = document.getElementById('weight-min');
    var weightMax = document.getElementById('weight-max');
    var applyBtn = document.getElementById('apply-filters-btn');
    var countEl = document.getElementById('results-count');
    var sortSelect = document.getElementById('sort-select');

    function activeCategories() {
      var checked = catBoxes.filter(function (b) { return b.checked; }).map(function (b) { return b.value; });
      return checked; // empty array = no category filter applied (show all)
    }
    function activeGrades() {
      var active = gradeChips.filter(function (c) { return c.classList.contains('active'); }).map(function (c) { return c.dataset.grade; });
      return active; // empty = show all grades
    }

    function applyFilters() {
      var cats = activeCategories();
      var grades = activeGrades();
      var certOnly = certBox ? certBox.checked : false;
      var minW = weightMin && weightMin.value !== '' ? parseFloat(weightMin.value) : null;
      var maxW = weightMax && weightMax.value !== '' ? parseFloat(weightMax.value) : null;
      var visible = 0;

      cards.forEach(function (card) {
        var cCat = card.dataset.category;
        var cGrade = card.dataset.grade;
        var cCert = card.dataset.certified === 'true';
        var cWeight = parseFloat(card.dataset.weight || '0');

        var pass = true;
        if (cats.length && cats.indexOf(cCat) === -1) pass = false;
        if (pass && grades.length && grades.indexOf(cGrade) === -1) pass = false;
        if (pass && certOnly && !cCert) pass = false;
        if (pass && minW !== null && cWeight < minW) pass = false;
        if (pass && maxW !== null && cWeight > maxW) pass = false;

        card.style.display = pass ? '' : 'none';
        if (pass) visible++;
      });

      if (countEl) {
        countEl.textContent = 'Showing ' + visible + ' of ' + cards.length + ' listings';
      }
    }

    function applySort() {
      if (!sortSelect) return;
      var val = sortSelect.value;
      var sorted = cards.slice();
      if (val === 'price-asc') sorted.sort(function (a, b) { return parseFloat(a.dataset.price) - parseFloat(b.dataset.price); });
      else if (val === 'price-desc') sorted.sort(function (a, b) { return parseFloat(b.dataset.price) - parseFloat(a.dataset.price); });
      else if (val === 'co2') sorted.sort(function (a, b) { return parseFloat(b.dataset.co2) - parseFloat(a.dataset.co2); });
      sorted.forEach(function (card) { grid.appendChild(card); });
    }

    catBoxes.forEach(function (b) { b.addEventListener('change', applyFilters); });
    gradeChips.forEach(function (c) {
      c.addEventListener('click', function () {
        c.classList.toggle('active');
        applyFilters();
      });
    });
    if (certBox) certBox.addEventListener('change', applyFilters);
    if (weightMin) weightMin.addEventListener('input', applyFilters);
    if (weightMax) weightMax.addEventListener('input', applyFilters);
    if (applyBtn) applyBtn.addEventListener('click', function (e) { e.preventDefault(); applyFilters(); });
    if (sortSelect) sortSelect.addEventListener('change', applySort);

    /* Pre-select category coming from a homepage link like marketplace.html?cat=plastic */
    var params = new URLSearchParams(window.location.search);
    var urlCat = params.get('cat');
    if (urlCat) {
      catBoxes.forEach(function (b) { b.checked = (b.value === urlCat); });
    }
    applyFilters();
  }

  /* =========================================================
     Registration: different fields for Individual vs Business
     (Individual -> FIN code from ID card / passport, Business -> VÖEN)
     ========================================================= */
  var acctRadios = document.querySelectorAll('input[name="acct"]');
  if (acctRadios.length) {
    var individualFields = document.querySelectorAll('[data-account-fields="individual"]');
    var businessFields = document.querySelectorAll('[data-account-fields="business"]');
    var nameLabel = document.getElementById('r-name-label');
    var nameInput = document.getElementById('r-name');
    function toggleAccountFields() {
      var isBusiness = document.querySelector('input[name="acct"]:checked').value === 'business';
      individualFields.forEach(function (el) {
        el.style.display = isBusiness ? 'none' : '';
        var input = el.querySelector('input');
        if (input) input.required = !isBusiness;
      });
      businessFields.forEach(function (el) {
        el.style.display = isBusiness ? '' : 'none';
        var input = el.querySelector('input');
        if (input) input.required = isBusiness;
      });
      if (nameLabel) nameLabel.textContent = isBusiness ? 'Company name' : 'Full name';
      if (nameInput) nameInput.placeholder = isBusiness ? 'e.g. GreenLoop Recycling LLC' : 'e.g. Ravil Abbasov';
    }
    acctRadios.forEach(function (r) { r.addEventListener('change', toggleAccountFields); });
    toggleAccountFields();
  }
});
