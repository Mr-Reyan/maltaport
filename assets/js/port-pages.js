// Port Pages JavaScript - Centralized for all 49 port pages

// --- Port List Data for Dropdown (49 Ports) ---
const ports = [
  "Ship Chandler Dar Es Salaam", "Ship Chandler Apapa", "Ship Chandler Onne Port Nigeria", "San Pedro Ship Supply Chandler",
  "Ship Chandler Abidjan", "Ship Chandler Freetown", "Monrovia Ship Chandler", "Ship Chandler Maputo", "Ship Chandler Matadi",
  "Ship Chandler Beira", "Ship Chandler Angola", "Ship Chandler Lobito", "ShipChandler Lome", "Shipchandler Lome",
  "Ivory Coast Ship Chandler", "Nouadhibou Ship Chandler", "Ship Chandler Luanda", "Ship Chandler Cape Verde",
  "Ship Chandler Gabon", "Ship Chandler Namibia", "Walvis Bay Ship Chandler", "Ship Chandler Gentil Port - Gabon",
  "Ship Chandler Takoradi", "Ship Chandler Congo", "Shipchandler Lagos", "Ship Chandler Port Ambon",
  "Shipchandler Pointe Noire", "Ship Chandler Djibouti", "Ship Chandler Cabinda", "Ship Chandler Tema",
  "Ship Chandler Mindelo", "Ship Chandler Casablanca", "Ship Chandler Safi Port Morocco", "Ship Chandler Kamsar",
  "Ship Chandler Soyo", "Dakar Ship Chandler", "Ship Chandler Senegal", "Ship Chandler Dakar", "Bissau Ship Chandler - Guinea",
  "Conakry Ship Chandler", "Ship Chandler Cotonou", "Douala Ship Chandler", "Liberia Ship Chandler", "Namibia Ship Chandler",
  "Ship Chandler Dakar", "Ship Chandler Mombasa", "Ship Chandlers Saldanha Bay", "Ship Chandler Owendo", "Ship Chandlers Richards Bay"
];

// Populate Mega Menu
const megaList = document.getElementById('mega-port-list');
if (megaList) {
  const third = Math.ceil(ports.length / 3);
  const col1 = ports.slice(0, third);
  const col2 = ports.slice(third, third * 2);
  const col3 = ports.slice(third * 2);

  function createPortGroup(title, items) {
    const group = document.createElement('div');
    group.className = 'mega-group';
    group.innerHTML = `<h4>${title}</h4>`;
    items.forEach(port => {
      const a = document.createElement('a');
      const isPortPage = window.location.pathname.includes('/ports/') || window.location.pathname.includes('\\ports\\');
      const prefix = isPortPage ? '' : 'ports/';
      a.href = `${prefix}${port.toLowerCase().replace(/[^a-z0-9]+/g, '-')}.html`;
      a.textContent = port;
      group.appendChild(a);
    });
    return group;
  }

  megaList.appendChild(createPortGroup('African Ports (A-D)', col1));
  megaList.appendChild(createPortGroup('African Ports (E-M)', col2));
  megaList.appendChild(createPortGroup('African Ports (N-Z)', col3));

  // Mega Menu Search
  const searchInput = document.getElementById('mega-port-search');
  if (searchInput) {
    searchInput.addEventListener('input', function (e) {
      const query = e.target.value.toLowerCase();
      const allLinks = megaList.querySelectorAll('a');
      const groups = megaList.querySelectorAll('.mega-group');

      allLinks.forEach(link => {
        const text = link.textContent.toLowerCase();
        link.style.display = text.includes(query) ? 'block' : 'none';
      });

      groups.forEach(group => {
        const hasVisible = Array.from(group.querySelectorAll('a')).some(a => a.style.display !== 'none');
        group.style.display = hasVisible ? 'block' : 'none';
      });
    });
  }
}

document.addEventListener('DOMContentLoaded', function () {

  // --- Sticky Header Effect ---
  const header = document.querySelector('.site-header');
  if (header) {
    window.addEventListener('scroll', () => {
      if (window.scrollY > 20) {
        header.classList.add('scrolled');
      } else {
        header.classList.remove('scrolled');
      }
    });
  }

  // --- Mobile Nav Toggle ---
  const navToggle = document.querySelector('.nav-toggle');
  const primaryNav = document.querySelector('.primary-nav');

  if (navToggle && primaryNav) {
    navToggle.addEventListener('click', function () {
      primaryNav.classList.toggle('open');
      this.classList.toggle('open');
    });
  }

  // --- Mobile Dropdown Toggles ---
  const dropdownItems = document.querySelectorAll('.nav-item.has-dropdown, .nav-item.has-mega');
  dropdownItems.forEach(item => {
    const link = item.querySelector('.nav-link');
    if (link) {
      link.addEventListener('click', function (e) {
        if (window.innerWidth <= 1024) {
          e.preventDefault();
          item.classList.toggle('open');
        }
      });
    }
  });

  // --- Scroll Reveal Animation ---
  const revealElements = document.querySelectorAll('.reveal');
  if ('IntersectionObserver' in window) {
    const revealObserver = new IntersectionObserver((entries) => {
      entries.forEach(entry => {
        if (entry.isIntersecting) {
          entry.target.classList.add('active');
          revealObserver.unobserve(entry.target);
        }
      });
    }, {
      threshold: 0.15,
      rootMargin: '0px 0px -50px 0px'
    });

    revealElements.forEach(el => revealObserver.observe(el));
  } else {
    revealElements.forEach(el => el.classList.add('active'));
  }

  // --- Table of Contents Active State Tracking ---
  const tocLinks = document.querySelectorAll('.toc-link');
  const headings = document.querySelectorAll('.article-content h2[id]');

  if (tocLinks.length > 0 && headings.length > 0) {
    window.addEventListener('scroll', () => {
      let current = '';

      headings.forEach(heading => {
        const headingTop = heading.offsetTop;
        if (window.scrollY >= headingTop - 150) {
          current = heading.getAttribute('id');
        }
      });

      tocLinks.forEach(link => {
        link.classList.remove('active');
        if (link.getAttribute('href') === `#${current}`) {
          link.classList.add('active');
        }
      });
    });
  }

});
