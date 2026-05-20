const CONFIG = {
  dataUrl: './data.json',
  dublinCenter: { lat: 53.3498, lng: -6.2603 },
  nearbyRadiusKm: 5,
  enableAutoRefresh: false,
  autoRefreshMs: 60_000
};

const STATUS_COLORS = {
  Scheduled: '#2563eb',
  Completed: '#16a34a',
  Rescheduled: '#f59e0b',
  'Pending Allocation': '#6b7280'
};

let map;
let infoWindow;
let allRecords = [];
let visibleRecords = [];
let markers = [];
let markerCluster;
let directionsService;
let directionsRenderer;

function initMap() {
  map = new google.maps.Map(document.getElementById('map'), {
    center: CONFIG.dublinCenter,
    zoom: 11,
    mapTypeControl: false,
    streetViewControl: false
  });

  infoWindow = new google.maps.InfoWindow();
  directionsService = new google.maps.DirectionsService();
  directionsRenderer = new google.maps.DirectionsRenderer({ map, suppressMarkers: false });

  bindEvents();
  refreshData();

  if (CONFIG.enableAutoRefresh) {
    setInterval(refreshData, CONFIG.autoRefreshMs);
  }
}

async function refreshData() {
  // In SPFx production, replace this fetch with SharePoint List REST/Microsoft Graph retrieval.
  const response = await fetch(`${CONFIG.dataUrl}?t=${Date.now()}`);
  allRecords = await response.json();
  populateFilterDropdowns(allRecords);
  populateRouteDropdowns(allRecords);
  applyFilters();
}

function bindEvents() {
  document.getElementById('searchInput').addEventListener('input', applyFilters);
  document.getElementById('statusFilter').addEventListener('change', applyFilters);
  document.getElementById('surveyorFilter').addEventListener('change', applyFilters);
  document.getElementById('clearFiltersBtn').addEventListener('click', clearFilters);
  document.getElementById('refreshBtn').addEventListener('click', refreshData);
  document.getElementById('showRouteBtn').addEventListener('click', showRoute);
  document.getElementById('clearRouteBtn').addEventListener('click', clearRoute);
}

function populateFilterDropdowns(records) {
  const statusSelect = document.getElementById('statusFilter');
  const surveyorSelect = document.getElementById('surveyorFilter');
  const statuses = [...new Set(records.map(r => r.status))].sort();
  const surveyors = [...new Set(records.map(r => r.surveyorName))].sort();

  statusSelect.innerHTML = '<option value="All">All Statuses</option>' + statuses.map(s => `<option value="${s}">${s}</option>`).join('');
  surveyorSelect.innerHTML = '<option value="All">All Surveyors</option>' + surveyors.map(s => `<option value="${s}">${s}</option>`).join('');
}

function populateRouteDropdowns(records) {
  const from = document.getElementById('routeFromSelect');
  const to = document.getElementById('routeToSelect');
  const options = records.map(r => `<option value="${r.id}">${r.id} - ${r.address}</option>`).join('');
  from.innerHTML = '<option value="">Select start record</option>' + options;
  to.innerHTML = '<option value="">Select end record</option>' + options;
}

function applyFilters() {
  const search = document.getElementById('searchInput').value.trim().toLowerCase();
  const status = document.getElementById('statusFilter').value;
  const surveyor = document.getElementById('surveyorFilter').value;

  visibleRecords = allRecords.filter(r => {
    const matchesSearch = !search || [r.id, r.eircode, r.address, r.tenantName].join(' ').toLowerCase().includes(search);
    const matchesStatus = status === 'All' || r.status === status;
    const matchesSurveyor = surveyor === 'All' || r.surveyorName === surveyor;
    return matchesSearch && matchesStatus && matchesSurveyor;
  });

  drawMarkers(visibleRecords);
  updateSummary();
}

function drawMarkers(records) {
  markers.forEach(m => m.setMap(null));
  if (markerCluster) markerCluster.clearMarkers();

  markers = records.map(record => {
    const marker = new google.maps.Marker({
      position: { lat: record.latitude, lng: record.longitude },
      map,
      title: `${record.id} - ${record.address}`,
      icon: createMarkerIcon(record.status, record.surveyorInitials)
    });

    marker.addListener('click', () => {
      showInfo(record, marker);
      showNearbyRecords(record);
    });

    return marker;
  });

  markerCluster = new markerClusterer.MarkerClusterer({ map, markers });
}

function createMarkerIcon(status, initials) {
  const fill = STATUS_COLORS[status] || '#334155';
  const label = status === 'Pending Allocation' ? '?' : (initials || '?');

  const svg = `
  <svg xmlns="http://www.w3.org/2000/svg" width="48" height="48" viewBox="0 0 48 48">
    <circle cx="24" cy="20" r="14" fill="${fill}" stroke="#0f172a" stroke-width="2"></circle>
    <path d="M24 46 L17 30 L31 30 Z" fill="${fill}" stroke="#0f172a" stroke-width="2"></path>
    <text x="24" y="24" text-anchor="middle" font-size="10" fill="#fff" font-family="Arial" font-weight="700">${label}</text>
  </svg>`;

  return {
    url: `data:image/svg+xml;charset=UTF-8,${encodeURIComponent(svg)}`,
    scaledSize: new google.maps.Size(36, 36),
    anchor: new google.maps.Point(18, 34)
  };
}

function showInfo(record, marker) {
  const content = `
    <div style="max-width:280px">
      <strong>Reference ID:</strong> ${record.id}<br>
      <strong>Eircode:</strong> ${record.eircode}<br>
      <strong>Address:</strong> ${record.address}<br>
      <strong>Tenant:</strong> ${record.tenantName}<br>
      <strong>Phone:</strong> ${record.tenantPhone}<br>
      <strong>Surveyor:</strong> ${record.surveyorName}<br>
      <strong>Appointment:</strong> ${new Date(record.appointmentDateTime).toLocaleString()}<br>
      <strong>Status:</strong> ${record.status}<br>
      <strong>Comments:</strong> ${record.comments}
    </div>`;

  infoWindow.setContent(content);
  infoWindow.open({ map, anchor: marker });
}

function updateSummary() {
  const total = allRecords.length;
  const visible = visibleRecords.length;
  const counts = ['Scheduled', 'Completed', 'Rescheduled', 'Pending Allocation']
    .map(s => `${s}: ${visibleRecords.filter(r => r.status === s).length}`);

  document.getElementById('summaryList').innerHTML = `
    <li>Total records: ${total}</li>
    <li>Visible records: ${visible}</li>
    ${counts.map(c => `<li>${c}</li>`).join('')}
  `;
}

function clearFilters() {
  document.getElementById('searchInput').value = '';
  document.getElementById('statusFilter').value = 'All';
  document.getElementById('surveyorFilter').value = 'All';
  applyFilters();
}

function showNearbyRecords(selected) {
  const nearby = allRecords
    .filter(r => r.id !== selected.id)
    .map(r => ({ ...r, distanceKm: haversineKm(selected.latitude, selected.longitude, r.latitude, r.longitude) }))
    .filter(r => r.distanceKm <= CONFIG.nearbyRadiusKm)
    .sort((a, b) => a.distanceKm - b.distanceKm)
    .slice(0, 5);

  const list = document.getElementById('nearbyList');
  list.innerHTML = nearby.length
    ? nearby.map(r => `<li><strong>${r.id}</strong> (${r.distanceKm.toFixed(2)} km) - ${r.address}</li>`).join('')
    : '<li>No nearby appointments within radius.</li>';
}

function haversineKm(lat1, lng1, lat2, lng2) {
  const toRad = d => (d * Math.PI) / 180;
  const R = 6371;
  const dLat = toRad(lat2 - lat1);
  const dLng = toRad(lng2 - lng1);
  const a = Math.sin(dLat / 2) ** 2 + Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLng / 2) ** 2;
  return 2 * R * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

function showRoute() {
  const fromId = document.getElementById('routeFromSelect').value;
  const toId = document.getElementById('routeToSelect').value;
  if (!fromId || !toId || fromId === toId) return;

  const from = allRecords.find(r => r.id === fromId);
  const to = allRecords.find(r => r.id === toId);

  directionsService.route(
    {
      origin: { lat: from.latitude, lng: from.longitude },
      destination: { lat: to.latitude, lng: to.longitude },
      travelMode: google.maps.TravelMode.DRIVING
    },
    (result, status) => {
      if (status === 'OK') {
        directionsRenderer.setDirections(result);
      } else {
        alert('Route request failed. Ensure Directions API is enabled for your key.');
      }
    }
  );
}

function clearRoute() {
  directionsRenderer.set('directions', null);
}

window.initMap = initMap;
