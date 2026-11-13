// GreenLoop — Cart & Buyer/Seller Request demo logic
// Uses localStorage so the whole flow works across pages without a backend.
// NOTE: for the buyer <-> seller demo to work, open the site through a local
// server (or Chrome's file:// works too) so localStorage persists between pages.

(function () {
  var CART_KEY = 'gl_cart';
  var REQ_KEY = 'gl_requests';
  var CURRENT_BUYER = 'Ravil Abbasov';

  /* ---------------- storage helpers ---------------- */
  function readJSON(key) {
    try {
      var raw = localStorage.getItem(key);
      return raw ? JSON.parse(raw) : [];
    } catch (e) { return []; }
  }
  function writeJSON(key, val) {
    try { localStorage.setItem(key, JSON.stringify(val)); } catch (e) { /* ignore */ }
  }

  /* ---------------- cart API ---------------- */
  function getCart() { return readJSON(CART_KEY); }
  function saveCart(arr) { writeJSON(CART_KEY, arr); updateCartBadge(); }

  function addToCart(item) {
    var cart = getCart();
    var existing = cart.find(function (c) { return c.id === item.id; });
    if (existing) {
      existing.weight = Math.round((existing.weight + item.weight) * 100) / 100;
    } else {
      cart.push(item);
    }
    saveCart(cart);
    return cart;
  }
  function removeFromCart(id) {
    var cart = getCart().filter(function (c) { return c.id !== id; });
    saveCart(cart);
    return cart;
  }
  function updateCartWeight(id, weight) {
    var cart = getCart();
    var item = cart.find(function (c) { return c.id === id; });
    if (item) item.weight = Math.max(0.1, weight);
    saveCart(cart);
    return cart;
  }
  function clearCart() { saveCart([]); }

  function updateCartBadge() {
    var count = getCart().length;
    document.querySelectorAll('#cart-badge').forEach(function (el) {
      el.textContent = count;
      el.style.display = count > 0 ? 'flex' : 'none';
    });
  }

  /* ---------------- requests (offers / cart orders) API ---------------- */
  function getRequests() { return readJSON(REQ_KEY); }
  function saveRequests(arr) { writeJSON(REQ_KEY, arr); }

  function addRequest(req) {
    var list = getRequests();
    req.id = 'REQ-' + Date.now().toString(36).toUpperCase() + Math.floor(Math.random() * 100);
    req.status = 'pending';
    req.createdAt = new Date().toISOString();
    req.buyer = req.buyer || CURRENT_BUYER;
    list.unshift(req);
    saveRequests(list);
    return req;
  }
  function updateRequestStatus(id, status) {
    var list = getRequests();
    var r = list.find(function (r) { return r.id === id; });
    if (r) r.status = status;
    saveRequests(list);
    return list;
  }

  function checkoutCart(message) {
    var cart = getCart();
    cart.forEach(function (item) {
      addRequest({
        type: 'cart',
        title: item.title,
        category: item.category,
        matClass: item.matClass,
        icon: item.icon,
        weight: item.weight,
        pricePerKg: item.price,
        total: Math.round(item.weight * item.price * 100) / 100,
        seller: item.seller,
        message: message || ''
      });
    });
    clearCart();
  }

  function timeAgo(iso) {
    var diff = Date.now() - new Date(iso).getTime();
    var mins = Math.floor(diff / 60000);
    if (mins < 1) return 'just now';
    if (mins < 60) return mins + ' min ago';
    var hrs = Math.floor(mins / 60);
    if (hrs < 24) return hrs + ' h ago';
    return Math.floor(hrs / 24) + ' d ago';
  }

  // Expose globally
  window.GL = {
    getCart: getCart, addToCart: addToCart, removeFromCart: removeFromCart,
    updateCartWeight: updateCartWeight, clearCart: clearCart, updateCartBadge: updateCartBadge,
    getRequests: getRequests, addRequest: addRequest, updateRequestStatus: updateRequestStatus,
    checkoutCart: checkoutCart, timeAgo: timeAgo, CURRENT_BUYER: CURRENT_BUYER
  };

  document.addEventListener('DOMContentLoaded', function () {
    updateCartBadge();

    /* Generic "Add to Cart" buttons: [data-add-cart] sits inside an element
       carrying the data-* fields (usually the .listing-card itself). */
    document.querySelectorAll('[data-add-cart]').forEach(function (btn) {
      btn.addEventListener('click', function (e) {
        e.preventDefault();
        e.stopPropagation();
        var card = btn.closest('[data-id]');
        if (!card) return;
        var d = card.dataset;
        GL.addToCart({
          id: d.id,
          title: d.title,
          category: d.category,
          matClass: d.matclass,
          icon: d.icon,
          price: parseFloat(d.price),
          weight: parseFloat(d.defaultWeight || d.weight || '1'),
          seller: d.seller
        });
        var original = btn.textContent;
        btn.textContent = '✓ Added';
        btn.classList.add('btn-primary');
        btn.classList.remove('btn-outline');
        setTimeout(function () {
          btn.textContent = original;
          btn.classList.remove('btn-primary');
          btn.classList.add('btn-outline');
        }, 1400);
      });
    });
  });
})();
