#!/usr/bin/env python3
"""
Auto-patch script: base.html mein sidebar update karta hai
Usage: python3 apply_sidebar_patch.py
"""
import re, os, shutil

BASE_HTML = os.path.join(os.path.dirname(__file__), 'templates', 'base.html')

if not os.path.exists(BASE_HTML):
    print("ERROR: templates/base.html nahi mila!")
    print("Manually BASE_HTML_SIDEBAR_PATCH.html dekh ke apply karo.")
    exit(1)

# Backup
shutil.copy(BASE_HTML, BASE_HTML + '.backup')
print(f"Backup bana diya: {BASE_HTML}.backup")

with open(BASE_HTML, 'r', encoding='utf-8') as f:
    content = f.read()

# ── Check if already patched ──
if 'roster-submenu' in content:
    print("Already patched hai! Koi change nahi hua.")
    exit(0)

# ── CSS to inject before </style> or </head> ──
SIDEBAR_CSS = """
/* ── Roster Dropdown Sidebar ── */
.roster-submenu {
  list-style:none; padding:4px 0; margin:0; display:none;
  background:rgba(255,255,255,0.06);
  border-left:3px solid #0d6efd;
  margin-left:14px; border-radius:0 6px 6px 0;
}
.roster-dropdown-item:hover .roster-submenu,
.roster-dropdown-item:focus-within .roster-submenu { display:block; }
.roster-dropdown-item:hover .roster-chevron { transform:rotate(180deg); }
.roster-sub-link {
  display:flex; align-items:center; gap:8px;
  padding:7px 14px 7px 12px;
  color:rgba(255,255,255,0.75); text-decoration:none;
  font-size:13px; transition:.15s; border-radius:0 6px 6px 0;
}
.roster-sub-link:hover { background:rgba(13,110,253,0.25); color:#fff; }
.roster-sub-link.active-sub { background:#0d6efd; color:#fff; font-weight:600; }
.rlink-icon { font-size:14px; width:20px; text-align:center; }
.roster-chevron { font-size:10px; transition:.2s; }
"""

# ── JS to inject before </body> ──
SIDEBAR_JS = """
<script>
(function(){
  function highlightRosterNav(){
    var p=new URLSearchParams(window.location.search),tab=p.get('tab')||'roster';
    var isRoster=window.location.pathname==='/roster';
    var menu=document.getElementById('rosterSubmenu');
    if(menu&&isRoster) menu.style.display='block';
    ['roster','rules','mapping'].forEach(function(t){
      var el=document.getElementById('rlink_'+t);
      if(!el)return;
      el.classList.toggle('active-sub', isRoster && tab===t);
    });
  }
  document.readyState==='loading'
    ? document.addEventListener('DOMContentLoaded',highlightRosterNav)
    : highlightRosterNav();
})();
</script>
"""

# ── New sidebar LI HTML ──
NEW_LI = """<li class="nav-item roster-dropdown-item" style="position:relative;">
          <a class="nav-link d-flex align-items-center gap-2" href="/roster?tab=roster">
            <i class="bi bi-calendar-week"></i>
            <span>Roster Management</span>
            <i class="bi bi-chevron-down ms-auto roster-chevron"></i>
          </a>
          <ul class="roster-submenu" id="rosterSubmenu">
            <li><a href="/roster?tab=roster"  class="roster-sub-link" id="rlink_roster"><span class="rlink-icon">&#128197;</span> Roster Management</a></li>
            <li><a href="/roster?tab=rules"   class="roster-sub-link" id="rlink_rules"><span class="rlink-icon">&#128203;</span> Roster Rules Template</a></li>
            <li><a href="/roster?tab=mapping" class="roster-sub-link" id="rlink_mapping"><span class="rlink-icon">&#128101;</span> Employee Mapping</a></li>
          </ul>
        </li>"""

# ── Pattern to find existing roster nav item ──
# Handles different possible formats in base.html
patterns = [
    # With icon
    r'<li[^>]*>\s*<a[^>]*href=["\']/?roster["\'][^>]*>.*?Weekly Roster.*?</a>\s*</li>',
    r'<li[^>]*>\s*<a[^>]*>.*?Weekly Roster.*?</a>\s*</li>',
    # Simple format
    r'<li>\s*<a[^>]*href=["\']/?roster["\'][^>]*>[^<]*Weekly Roster[^<]*</a>\s*</li>',
]

replaced = False
for pat in patterns:
    match = re.search(pat, content, re.DOTALL | re.IGNORECASE)
    if match:
        content = content[:match.start()] + NEW_LI + content[match.end():]
        print(f"✅ Sidebar li replaced! (pattern: {pat[:50]}...)")
        replaced = True
        break

if not replaced:
    print("⚠️  Automatic replacement nahi ho saka.")
    print("   Manually base.html mein 'Weekly Roster' wali <li> dhundho")
    print("   aur BASE_HTML_SIDEBAR_PATCH.html se replace karo.")
    print("")
    print("   NEW <li> code yahan se copy karo:")
    print(NEW_LI)

# Inject CSS
if '</head>' in content:
    content = content.replace('</head>', f'<style>{SIDEBAR_CSS}</style>\n</head>', 1)
    print("✅ CSS injected before </head>")
elif '</style>' in content:
    # inject before last </style>
    idx = content.rfind('</style>')
    content = content[:idx] + SIDEBAR_CSS + content[idx:]
    print("✅ CSS injected in existing <style>")

# Inject JS
if '</body>' in content:
    content = content.replace('</body>', SIDEBAR_JS + '\n</body>', 1)
    print("✅ JS injected before </body>")

with open(BASE_HTML, 'w', encoding='utf-8') as f:
    f.write(content)

print("")
print("=" * 50)
print("PATCH COMPLETE! Server restart karo.")
print("Backup hai: templates/base.html.backup")
print("=" * 50)
