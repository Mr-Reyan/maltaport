import os
import re
import openpyxl

# Configuration Paths
REPO_ROOT = os.path.dirname(os.path.abspath(__file__))
PORTS_DIR = os.path.join(REPO_ROOT, "ports")
CSS_FILE = os.path.join(REPO_ROOT, "assets", "css", "port-pages.css")
JS_FILE = os.path.join(REPO_ROOT, "assets", "js", "port-pages.js")
EXCEL_PATH = r"D:\FalconXoft Ship Chandler\Website Data\Ship_Chandler_SEO_49_Pages_FULL.xlsx"

PORT_PAGE_TEMPLATE = """<!DOCTYPE html>
<html lang="en">

<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
  <title>__META_TITLE__</title>
  <meta name="description"
    content="__META_DESC__">
  <meta name="keywords"
    content="__KEYWORDS__">
  <meta name="author" content="SHIP CHANDLER MALTA CO., LTD">
  <link rel="canonical" href="https://shipchandlermalta.com/ports/__SLUG__">

  <!-- Open Graph Meta Tags -->
  <meta property="og:title" content="__META_TITLE__">
  <meta property="og:description"
    content="__META_DESC__">
  <meta property="og:type" content="website">
  <meta property="og:url" content="https://shipchandlermalta.com/ports/__SLUG__">
  <meta property="og:site_name" content="Ship Chandler Malta">

  <!-- Twitter Card Meta Tags -->
  <meta name="twitter:card" content="summary_large_image">
  <meta name="twitter:title" content="__META_TITLE__">
  <meta name="twitter:description"
    content="__META_DESC__">

  <!-- Google Fonts (Elegant Serif + Clean Sans) -->
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link
    href="https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=Manrope:wght@300;400;500;600;700&display=swap"
    rel="stylesheet">

  <!-- Font Awesome -->
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css">

  <!-- Favicon & App Icons -->
  <link rel="icon" type="image/png" href="../favicon.png">
  <link rel="shortcut icon" href="../favicon.ico">
  <link rel="apple-touch-icon" href="../favicon.png">

  <!-- Centralized Port Pages CSS -->
  <link rel="stylesheet" href="../assets/css/port-pages.css">
</head>

<body>

  <!-- HEADER / NAVBAR -->
  <header class="site-header" id="site-header">
    <div class="container header-inner">
      <a class="logo" href="../index.html">
        <img src="../assets/imgs/logo.png" alt="ShipChandler Malta Logo" width="44" height="44">
        <span class="logo-text">
          <strong>ShipChandler</strong>
          <span>Malta</span>
        </span>
      </a>

      <button class="nav-toggle" type="button" aria-label="Toggle navigation">
        <span></span><span></span><span></span>
      </button>

      <nav class="primary-nav" id="primary-nav">
        <ul class="nav-list">
          <li class="nav-item"><a class="nav-link" href="../index.html">Home</a></li>
          <li class="nav-item"><a class="nav-link" href="../about.html">About</a></li>

          <li class="nav-item has-dropdown">
            <a class="nav-link" href="../services.html">Services <span class="nav-arrow"><svg viewBox="0 0 24 24"
                  width="12" height="12" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"
                  stroke-linejoin="round">
                  <polyline points="6 9 12 15 18 9"></polyline>
                </svg></span></a>
            <div class="dropdown-menu">
              <a href="../services.html#provisions">Food Provisions Supply</a>
              <a href="../services.html#bonded">Bonded Stores</a>
              <a href="../services.html#deck">Deck Stores</a>
              <a href="../services.html#engine">Engine Stores</a>
              <a href="../services.html#cabin">Cabin & Galley</a>
              <a href="../services.html#safety">Safety & PPE</a>
              <a href="../services.html#electrical">Electrical Stores</a>
              <a href="../services.html#chemicals">Marine Chemicals</a>
              <a href="../services.html#welding">Welding Equipment</a>
              <a href="../services.html#spare-parts">Spare Parts</a>
              <a href="../services.html#publications">Nautical Charts</a>
              <a href="../services.html#tools">Tools & Hardware</a>
              <a href="../services.html#urgent">Urgent Supply</a>
            </div>
          </li>

          <li class="nav-item has-mega active">
            <a class="nav-link" href="javascript:void(0)">Ports <span class="nav-arrow"><svg viewBox="0 0 24 24"
                  width="12" height="12" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"
                  stroke-linejoin="round">
                  <polyline points="6 9 12 15 18 9"></polyline>
                </svg></span></a>
            <div class="mega-menu">
              <div class="mega-header">
                <div class="mega-header-title">Global Port Network</div>
                <div class="mega-search">
                  <i class="fas fa-search" style="color:var(--silver);"></i>
                  <input type="text" id="mega-port-search" placeholder="Search 49 ports...">
                </div>
              </div>
              <div class="mega-grid" id="mega-port-list">
                <!-- Populated via JS -->
              </div>
            </div>
          </li>

          <li class="nav-item"><a class="nav-link" href="../contact.html">Contact</a></li>
        </ul>
        <a class="btn btn-primary" style="padding: 14px 32px; font-size: 0.82rem;" href="../request-quote.html">Get
          Quote</a>
      </nav>
    </div>
  </header>

  <main>
    <!-- SUBPAGE HERO -->
    <section class="subpage-hero">
      <div class="subpage-hero-bg">
        <img src="../assets/imgs/port-hero.jpg" alt="__PAGE_NAME__ Background">
      </div>
      <div class="subpage-hero-overlay"></div>

      <div class="subpage-hero-wrapper">
        <ul class="breadcrumb">
          <li><a href="../index.html">Home</a></li>
          <li><span>/</span></li>
          <li>__PAGE_NAME__</li>
        </ul>
        <span class="eyebrow" style="justify-content: center; display: inline-flex;">__EYEBROW__</span>
        <h1 class="subpage-hero-title">__PAGE_NAME__</h1>
        <p class="subpage-hero-desc">__META_DESC__</p>
      </div>
    </section>

    <!-- ARTICLE SECTION -->
    <section class="article-section">
      <div class="container article-grid">

        <!-- Sticky Table of Contents (Left) -->
        <aside class="toc-sidebar reveal">
          <h4 class="toc-title">Table of Contents</h4>
          <ul class="toc-list">
            <li><a href="#professional-services" class="toc-link">Professional Ship Supply</a></li>
            <li><a href="#provisions" class="toc-link">Ship Provisions</a></li>
            <li><a href="#bonded" class="toc-link">Bonded Stores</a></li>
            <li><a href="#deck" class="toc-link">Deck Stores</a></li>
            <li><a href="#engine" class="toc-link">Engine Stores</a></li>
            <li><a href="#cabin" class="toc-link">Cabin & Galley</a></li>
            <li><a href="#safety" class="toc-link">Safety Equipment</a></li>
            <li><a href="#electrical" class="toc-link">Electrical Stores</a></li>
            <li><a href="#chemicals" class="toc-link">Marine Chemicals</a></li>
            <li><a href="#welding" class="toc-link">Welding Equipment</a></li>
            <li><a href="#spare-parts" class="toc-link">Spare Parts</a></li>
            <li><a href="#publications" class="toc-link">Nautical Charts</a></li>
            <li><a href="#tools" class="toc-link">Tools & Hardware</a></li>
            <li><a href="#urgent" class="toc-link">Urgent Supply</a></li>
            <li><a href="#vessels" class="toc-link">Vessels Supported</a></li>
            <li><a href="#rfq" class="toc-link">How to Send RFQ</a></li>
            <li><a href="#why-choose" class="toc-link">Why Choose Us</a></li>
          </ul>
        </aside>

        <!-- Main Article Content (Right) -->
        <article class="article-content">
__ARTICLE_HTML__
        </article>
      </div>
    </section>
  </main>

  <!-- FOOTER -->
  <footer class="site-footer">
    <div class="container">
      <div class="footer-grid">
        <div class="footer-brand">
          <a class="logo" href="../index.html">
            <img src="../assets/imgs/logo.png" alt="Logo" width="44" height="44">
            <span class="logo-text">
              <strong>ShipChandler</strong>
              <span style="color: var(--crimson);">Malta</span>
            </span>
          </a>
          <p>SHIP CHANDLER MALTA CO., LTD is a licensed marine supplier providing food provisions, food
            stores, bond store supply, drinks, snacks, and emergency vessel supply services across Port
            Marsaxlokk, Malta.</p>
        </div>

        <div class="footer-col">
          <h4>Quick Links</h4>
          <ul class="footer-links">
            <li><a href="../index.html">Home</a></li>
            <li><a href="../about.html">About Us</a></li>
            <li><a href="../services.html">Services</a></li>
            <li><a href="../contact.html">Contact</a></li>
            <li><a href="../request-quote.html">Request a Quote</a></li>
          </ul>
        </div>

        <div class="footer-col">
          <h4>Services</h4>
          <ul class="footer-links">
            <li><a href="../services.html#provisions">Provisions Supply</a></li>
            <li><a href="../services.html#bonded">Bonded Stores</a></li>
            <li><a href="../services.html#deck">Deck Stores</a></li>
            <li><a href="../services.html#engine">Engine Stores</a></li>
            <li><a href="../services.html#urgent">Urgent Supply</a></li>
          </ul>
        </div>

        <div class="footer-col">
          <h4>Contact Us</h4>
          <ul class="footer-contact">
            <li><i class="fas fa-map-marker-alt"></i> Port of Marsaxlokk, Xatt il-Mollijiet, MRS 1152 Marsaxlokk, Malta</li>
            <li><i class="fas fa-envelope"></i> <a href="mailto:sales@shipchandlermalta.com"
                style="color:inherit;">sales@shipchandlermalta.com</a></li>
            <li><i class="fas fa-user-tie"></i> Matthew Borg</li>
            <li>
              <a class="hero-whatsapp-bar" href="https://wa.me/6590190005" target="_blank" rel="noopener noreferrer"
                aria-label="Direct WhatsApp contact (+65) 9019 0005">
                <span class="hero-whatsapp-icon" aria-hidden="true">
                  <svg viewBox="0 0 24 24" width="20" height="20" fill="currentColor">
                    <path
                      d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z" />
                  </svg>
                </span>
                <span class="hero-whatsapp-tag">WhatsApp:</span>
                <span class="hero-whatsapp-num">(+65) 9019 0005</span>
                <span class="hero-whatsapp-arrow" aria-hidden="true">&rarr;</span>
              </a>
            </li>
          </ul>
        </div>
      </div>

      <div class="footer-bottom">
        <div>&copy; 2026 SHIP CHANDLER MALTA CO., LTD. All rights reserved.</div>
      </div>
    </div>
  </footer>

  <!-- FLOATING WHATSAPP -->
  <a href="https://wa.me/6590190005" class="whatsapp-float" target="_blank" rel="noopener noreferrer"
    aria-label="WhatsApp Contact (+65) 9019 0005">
    <i class="fab fa-whatsapp"></i>
  </a>

  <!-- Centralized Port Pages JavaScript -->
  <script src="../assets/js/port-pages.js"></script>
</body>

</html>"""


def get_eyebrow(page_name):
    """
    Derives clean editorial eyebrow matching Bissau's 'Guinea Port Network'
    """
    raw = str(page_name).strip()
    if ' - ' in raw:
        loc = raw.split(' - ')[-1].strip()
        return f"{loc} Port Network"
    if 'Onne' in raw:
        return "Nigeria Port Network"
    if 'Safi' in raw:
        return "Morocco Port Network"

    clean = re.sub(r'(?i)\bship\s*chandlers?\b', '', raw)
    clean = re.sub(r'(?i)\bshipchandler\b', '', clean)
    clean = re.sub(r'(?i)\bship\s*supply\s*chandler\b', '', clean)
    clean = clean.strip()
    return f"{clean} Port Network"


def get_h2_id(title):
    """
    Maps section headings to anchor IDs matching the Table of Contents
    """
    t = title.lower()
    if 'professional ship supply' in t:
        return 'professional-services'
    if 'provisions' in t:
        return 'provisions'
    if 'bonded' in t:
        return 'bonded'
    if 'deck stores' in t:
        return 'deck'
    if 'engine stores' in t:
        return 'engine'
    if 'cabin' in t and 'galley' in t:
        return 'cabin'
    if 'safety' in t:
        return 'safety'
    if 'electrical' in t:
        return 'electrical'
    if 'chemical' in t:
        return 'chemicals'
    if 'welding' in t:
        return 'welding'
    if 'spare parts' in t:
        return 'spare-parts'
    if 'nautical' in t or 'publication' in t:
        return 'publications'
    if 'tools' in t:
        return 'tools'
    if 'technical and repair' in t:
        return 'technical-support'
    if 'urgent' in t:
        return 'urgent'
    if 'vessels supported' in t:
        return 'vessels'
    if 'rfq' in t:
        return 'rfq'
    if 'why choose' in t:
        return 'why-choose'
    return None


def render_article_html(desc_text):
    """
    Parses the Markdown description from Excel into semantic HTML
    matching the exact structure, tags, and indentation of bissau-ship-chandler-guinea.html
    """
    blocks = [b.strip() for b in desc_text.split('\n\n') if b.strip()]
    lines = []

    for b in blocks:
        if b.startswith('# '):
            # Skip top-level H1 (rendered in Subpage Hero)
            continue
        elif b.startswith('## '):
            heading_text = b[3:].strip()
            heading_text = heading_text.replace('\u2013', '–')
            h2_id = get_h2_id(heading_text)
            if h2_id:
                lines.append(f'          <h2 id="{h2_id}">{heading_text}</h2>')
            else:
                lines.append(f'          <h2>{heading_text}</h2>')
        elif b.startswith('- '):
            lines.append('          <ul>')
            for item in b.split('\n'):
                item = item.strip()
                if not item:
                    continue
                val = re.sub(r'^\s*-\s*', '', item)
                val = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', val)
                lines.append(f'            <li>{val}</li>')
            lines.append('          </ul>')
        else:
            # Inline bold markdown replacement & paragraph formatting
            content = b.replace('\n', ' ')
            content = re.sub(r'\s+', ' ', content)
            content = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', content)
            lines.append(f'          <p>{content}</p>')

    return '\n'.join(lines)


def generate_keywords(page_name):
    """
    Generates rich, tailored SEO keywords for the port page
    """
    raw = str(page_name).strip()
    clean_loc = re.sub(r'(?i)\bship\s*chandlers?\b', '', raw)
    clean_loc = re.sub(r'(?i)\bshipchandler\b', '', clean_loc)
    clean_loc = re.sub(r'(?i)\bship\s*supply\s*chandler\b', '', clean_loc)
    clean_loc = re.sub(r'\s*-\s*', ' ', clean_loc)
    clean_loc = re.sub(r'\s+', ' ', clean_loc).strip()

    kw_list = [
        raw.lower(),
        f"ship chandler {clean_loc.lower()}",
        f"ship supply {clean_loc.lower()}",
        f"port {clean_loc.lower()} vessel supply",
        "malta ship chandler",
        "provisions",
        "marine stores",
        "bonded stores",
        "deck stores",
        "engine stores",
        "vessel supply",
        "marine procurement"
    ]
    return ", ".join(kw_list)


def build_port_page_html(page_name, meta_title, meta_desc, full_desc, slug):
    """
    Builds the full HTML document for a port page matching Bissau line-for-line
    using centralized port-pages.css, centralized port-pages.js, and only the 4 project image assets:
      - ../favicon.ico
      - ../favicon.png
      - ../assets/imgs/logo.png
      - ../assets/imgs/port-hero.jpg
    """
    eyebrow = get_eyebrow(page_name)
    keywords = generate_keywords(page_name)
    article_html = render_article_html(full_desc)

    html = PORT_PAGE_TEMPLATE
    html = html.replace('__META_TITLE__', str(meta_title or ''))
    html = html.replace('__META_DESC__', str(meta_desc or ''))
    html = html.replace('__KEYWORDS__', keywords)
    html = html.replace('__PAGE_NAME__', str(page_name or ''))
    html = html.replace('__SLUG__', slug)
    html = html.replace('__EYEBROW__', eyebrow)
    html = html.replace('__ARTICLE_HTML__', article_html)

    return html


def main():
    print(f"Reading Excel data from: {EXCEL_PATH}")
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    sheet = wb['SEO Pages']

    rows = list(sheet.iter_rows(values_only=True))[1:]  # Skip header row
    print(f"Found {len(rows)} port rows in Excel.")

    os.makedirs(PORTS_DIR, exist_ok=True)

    generated_count = 0
    unique_slugs = set()

    for row in rows:
        num, page_name, h1, meta_title, meta_desc, full_desc = row[:6]
        if not page_name:
            continue

        slug = re.sub(r'[^a-z0-9]+', '-', str(page_name).lower()).strip('-') + '.html'
        out_path = os.path.join(PORTS_DIR, slug)

        html_content = build_port_page_html(
            page_name=page_name,
            meta_title=meta_title,
            meta_desc=meta_desc,
            full_desc=full_desc,
            slug=slug
        )

        with open(out_path, 'w', encoding='utf-8') as f:
            f.write(html_content)

        generated_count += 1
        unique_slugs.add(slug)
        print(f"[{num:02d}/49] Generated: {slug} ({page_name})")

    print("\n" + "=" * 60)
    print(f"SUCCESS: Processed {generated_count} rows across {len(unique_slugs)} unique port page files.")
    print("Centralized CSS: assets/css/port-pages.css")
    print("Centralized JS: assets/js/port-pages.js")
    print("Configured strictly to the 4 existing images:")
    print("  1. ../favicon.ico")
    print("  2. ../favicon.png")
    print("  3. ../assets/imgs/logo.png")
    print("  4. ../assets/imgs/port-hero.jpg")
    print("=" * 60)


if __name__ == "__main__":
    main()
