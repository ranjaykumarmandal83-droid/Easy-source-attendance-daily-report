"""
Run this script on server to patch my_attendance.html and settings.html
Usage: python3 patch_attendance_location.py
"""
import os, re

BASE = os.path.dirname(os.path.abspath(__file__))
TMPL = os.path.join(BASE, 'templates')

# ════════════════════════════════════════════════════════════════
# PATCH 1 — my_attendance.html: Add time + location to mark button
# ════════════════════════════════════════════════════════════════
ma_path = os.path.join(TMPL, 'my_attendance.html')
if os.path.exists(ma_path):
    with open(ma_path, 'r', encoding='utf-8') as f:
        ma = f.read()

    # JS to inject — replaces or appends before </body>
    location_js = """
<style>
.loc-badge{display:inline-flex;align-items:center;gap:5px;font-size:12px;padding:3px 10px;border-radius:20px;font-weight:600}
.loc-in{background:rgba(0,200,150,.15);color:#00c896;border:1px solid rgba(0,200,150,.3)}
.loc-out{background:rgba(239,68,68,.15);color:#ef4444;border:1px solid rgba(239,68,68,.3)}
.loc-checking{background:rgba(255,193,7,.15);color:#ffc107;border:1px solid rgba(255,193,7,.3)}
.loc-unknown{background:rgba(150,150,150,.15);color:#999;border:1px solid rgba(150,150,150,.3)}
#locPanel{background:#1c2333;border:1px solid rgba(255,255,255,.08);border-radius:12px;padding:14px;margin-bottom:16px}
</style>

<script>
// Override / wrap existing mark function with location + time
(function(){
  var _geoData = null;
  var _geoStatus = 'idle';

  function getLocation(cb){
    _geoStatus = 'checking';
    updateLocUI();
    if(!navigator.geolocation){
      _geoStatus = 'unsupported';
      updateLocUI();
      cb(null);
      return;
    }
    navigator.geolocation.getCurrentPosition(
      function(pos){
        _geoData = {latitude: pos.coords.latitude, longitude: pos.coords.longitude, accuracy: pos.coords.accuracy};
        _geoStatus = 'got';
        updateLocUI();
        // Reverse geocode (optional - use nominatim)
        fetch('https://nominatim.openstreetmap.org/reverse?lat='+pos.coords.latitude+'&lon='+pos.coords.longitude+'&format=json')
          .then(r=>r.json()).then(d=>{
            if(d && d.display_name) _geoData.address = d.display_name.substring(0,120);
            updateLocUI();
          }).catch(()=>{});
        cb(_geoData);
      },
      function(err){
        _geoStatus = 'denied';
        updateLocUI();
        cb(null);
      },
      {enableHighAccuracy:true, timeout:10000, maximumAge:60000}
    );
  }

  function updateLocUI(){
    var el = document.getElementById('locStatusBadge');
    var mapEl = document.getElementById('locMapLink');
    if(!el) return;
    var html = '', mapHtml = '';
    if(_geoStatus==='idle')      html = '<span class="loc-badge loc-unknown">📍 Location: Click karo to get</span>';
    if(_geoStatus==='checking')  html = '<span class="loc-badge loc-checking">⏳ Location fetch ho raha hai...</span>';
    if(_geoStatus==='denied')    html = '<span class="loc-badge loc-out">❌ Location access denied — Settings se allow karo</span>';
    if(_geoStatus==='unsupported') html = '<span class="loc-badge loc-out">❌ Browser location support nahi karta</span>';
    if(_geoStatus==='got' && _geoData){
      html = '<span class="loc-badge loc-in">✅ Location mila</span>';
      if(_geoData.address) html += '<div style="font-size:11px;color:#8b949e;margin-top:4px">📍 '+_geoData.address+'</div>';
      html += '<div style="font-size:11px;color:#8b949e;margin-top:2px">🎯 Accuracy: ~'+Math.round(_geoData.accuracy||0)+'m</div>';
      mapHtml = '<a href="https://maps.google.com?q='+_geoData.latitude+','+_geoData.longitude+'" target="_blank" style="font-size:11px;color:#3d5afe;margin-top:4px;display:inline-block">🗺 Map pe dekho</a>';
    }
    el.innerHTML = html;
    if(mapEl) mapEl.innerHTML = mapHtml;
  }

  // Intercept all fetch calls to /api/my/mark to inject location + time
  var _origFetch = window.fetch;
  window.fetch = function(url, opts){
    if(typeof url === 'string' && url.includes('/api/my/mark') && opts && opts.body){
      try{
        var body = JSON.parse(opts.body);
        body.check_in_time = new Date().toTimeString().substring(0,8);
        if(_geoData){
          body.latitude  = _geoData.latitude;
          body.longitude = _geoData.longitude;
          body.address   = _geoData.address || '';
        }
        opts.body = JSON.stringify(body);
      }catch(e){}
    }
    return _origFetch.apply(this, arguments);
  };

  // Auto-get location on page load
  document.addEventListener('DOMContentLoaded', function(){
    // Insert location panel before the mark attendance form
    var form = document.querySelector('form') || document.querySelector('[id*="mark"]') || document.querySelector('.card');
    if(form){
      var panel = document.createElement('div');
      panel.id = 'locPanel';
      panel.innerHTML = `
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px">
          <b style="font-size:13px;color:#e6edf3">📍 Location & Time</b>
          <span style="font-size:12px;color:#8b949e" id="locClock"></span>
        </div>
        <div id="locStatusBadge"><span class="loc-badge loc-checking">⏳ Location fetch ho raha hai...</span></div>
        <div id="locMapLink"></div>
      `;
      form.parentNode.insertBefore(panel, form);
    }

    // Live clock
    function updateClock(){
      var el = document.getElementById('locClock');
      if(el) el.textContent = '🕐 ' + new Date().toLocaleTimeString('en-IN');
    }
    updateClock();
    setInterval(updateClock, 1000);

    // Get location immediately
    getLocation(function(geo){
      if(!geo) console.log('Location not available');
    });
  });
})();
</script>
"""

    if 'locPanel' not in ma:
        # Insert before </body>
        if '</body>' in ma:
            ma = ma.replace('</body>', location_js + '\n</body>')
        else:
            ma += location_js
        with open(ma_path, 'w', encoding='utf-8') as f:
            f.write(ma)
        print("✅ my_attendance.html patched!")
    else:
        print("⚠️  my_attendance.html already patched")
else:
    print(f"❌ my_attendance.html not found at {ma_path}")

# ════════════════════════════════════════════════════════════════
# PATCH 2 — settings.html: Add Location Range section
# ════════════════════════════════════════════════════════════════
st_path = os.path.join(TMPL, 'settings.html')
if os.path.exists(st_path):
    with open(st_path, 'r', encoding='utf-8') as f:
        st = f.read()

    location_section = """
<!-- ═══════ LOCATION RANGE SETTINGS ═══════ -->
<div class="card shadow-sm mb-4" id="locationSettingsCard">
  <div class="card-header bg-success text-white d-flex justify-content-between align-items-center">
    <h5 class="mb-0">📍 Attendance Location Range</h5>
    <button class="btn btn-light btn-sm" onclick="addLocationRow()">+ Add Office</button>
  </div>
  <div class="card-body">
    <p class="text-muted small mb-3">
      Employees ke attendance mark karne ke liye allowed location range set karo.
      Agar employee is range ke bahar se mark kare to "Out of Range" show hoga.
    </p>
    <div class="table-responsive">
      <table class="table table-bordered table-sm" id="locSettingsTable">
        <thead class="table-dark">
          <tr>
            <th>Office Name</th>
            <th>Latitude</th>
            <th>Longitude</th>
            <th>Radius (meters)</th>
            <th>Active</th>
            <th>Action</th>
          </tr>
        </thead>
        <tbody id="locSettingsBody">
          {% for loc in loc_settings %}
          <tr id="locrow_{{ loc.id }}">
            <td><input type="text" class="form-control form-control-sm" id="loc_name_{{ loc.id }}" value="{{ loc.office_name }}"></td>
            <td><input type="number" step="0.000001" class="form-control form-control-sm" id="loc_lat_{{ loc.id }}" value="{{ loc.latitude }}"></td>
            <td><input type="number" step="0.000001" class="form-control form-control-sm" id="loc_lng_{{ loc.id }}" value="{{ loc.longitude }}"></td>
            <td><input type="number" class="form-control form-control-sm" id="loc_rad_{{ loc.id }}" value="{{ loc.radius_meters }}"></td>
            <td class="text-center">
              <input type="checkbox" id="loc_active_{{ loc.id }}" {% if loc.is_active %}checked{% endif %}>
            </td>
            <td>
              <button class="btn btn-primary btn-sm me-1" onclick="saveLocationRow({{ loc.id }})">Save</button>
              <button class="btn btn-sm" style="background:#fee2e2;color:#dc2626;border:1px solid #fca5a5" onclick="deleteLocationRow({{ loc.id }})">Delete</button>
              <a href="https://maps.google.com?q={{ loc.latitude }},{{ loc.longitude }}" target="_blank" class="btn btn-outline-secondary btn-sm">🗺</a>
            </td>
          </tr>
          {% endfor %}
        </tbody>
      </table>
    </div>
    <div class="alert alert-info small mt-2 mb-0">
      💡 <b>Tip:</b> Google Maps pe jaao → apni office location click karo → right-click → coordinates copy karo
    </div>
  </div>
</div>

<script>
function saveLocationRow(lid){
  var data = {
    id: lid,
    office_name: document.getElementById('loc_name_'+lid).value,
    latitude:    parseFloat(document.getElementById('loc_lat_'+lid).value),
    longitude:   parseFloat(document.getElementById('loc_lng_'+lid).value),
    radius_meters: parseInt(document.getElementById('loc_rad_'+lid).value)||200,
    is_active:   document.getElementById('loc_active_'+lid).checked ? 1 : 0
  };
  if(!data.office_name||!data.latitude||!data.longitude){alert('Sab fields fill karo!');return;}
  fetch('/api/location/settings/save',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(data)})
    .then(r=>r.json()).then(d=>{
      if(d.success) alert('✅ Saved!');
      else alert('Error: '+(d.error||''));
    });
}
function deleteLocationRow(lid){
  if(!confirm('Delete this location?')) return;
  fetch('/api/location/settings/delete',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({id:lid})})
    .then(r=>r.json()).then(d=>{
      if(d.success){ document.getElementById('locrow_'+lid).remove(); }
    });
}
var _newLocCount = 1000;
function addLocationRow(){
  var lid = _newLocCount++;
  var tr = document.createElement('tr');
  tr.id = 'locrow_'+lid;
  tr.innerHTML = `
    <td><input type="text" class="form-control form-control-sm" id="loc_name_${lid}" placeholder="Office name"></td>
    <td><input type="number" step="0.000001" class="form-control form-control-sm" id="loc_lat_${lid}" placeholder="28.6139"></td>
    <td><input type="number" step="0.000001" class="form-control form-control-sm" id="loc_lng_${lid}" placeholder="77.2090"></td>
    <td><input type="number" class="form-control form-control-sm" id="loc_rad_${lid}" value="200"></td>
    <td class="text-center"><input type="checkbox" id="loc_active_${lid}" checked></td>
    <td>
      <button class="btn btn-primary btn-sm me-1" onclick="saveNewLocation(${lid})">Save</button>
      <button class="btn btn-sm" style="background:#fee2e2;color:#dc2626" onclick="this.closest('tr').remove()">Remove</button>
    </td>`;
  document.getElementById('locSettingsBody').appendChild(tr);
}
function saveNewLocation(lid){
  var data = {
    office_name:   document.getElementById('loc_name_'+lid).value,
    latitude:      parseFloat(document.getElementById('loc_lat_'+lid).value),
    longitude:     parseFloat(document.getElementById('loc_lng_'+lid).value),
    radius_meters: parseInt(document.getElementById('loc_rad_'+lid).value)||200,
    is_active:     document.getElementById('loc_active_'+lid).checked ? 1 : 0
  };
  if(!data.office_name||!data.latitude||!data.longitude){alert('Sab fields fill karo!');return;}
  fetch('/api/location/settings/save',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(data)})
    .then(r=>r.json()).then(d=>{
      if(d.success){ alert('✅ Saved! Page reload karo.'); location.reload(); }
      else alert('Error: '+(d.error||''));
    });
}
</script>
"""

    if 'locationSettingsCard' not in st:
        # Insert before </main> or before last </div> of main content
        if '</main>' in st:
            st = st.replace('</main>', location_section + '\n</main>', 1)
        elif '{% endblock %}' in st:
            st = st.replace('{% endblock %}', location_section + '\n{% endblock %}', 1)
        else:
            st += location_section
        with open(st_path, 'w', encoding='utf-8') as f:
            f.write(st)
        print("✅ settings.html patched!")
    else:
        print("⚠️  settings.html already patched")
else:
    print(f"❌ settings.html not found at {st_path}")

print("\n✅ All patches done! Server restart karo.")
