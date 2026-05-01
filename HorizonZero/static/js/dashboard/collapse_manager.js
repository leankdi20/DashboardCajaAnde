document.addEventListener('DOMContentLoaded', function () {
 
  // Buscar todos los paneles registrados con data-collapse-panel
  var paneles = document.querySelectorAll('[data-collapse-panel]');
 
  paneles.forEach(function (panel) {
    var chevronId = panel.dataset.collapseChevron;
    var chevron   = chevronId ? document.getElementById(chevronId) : null;
    var hidden    = panel.dataset.collapseDefault === 'hide';
 
    if (hidden) {
      panel.classList.add('hidden');
      if (chevron) chevron.style.transform = 'rotate(-90deg)';
    } else {
      panel.classList.remove('hidden');
      if (chevron) chevron.style.transform = 'rotate(0deg)';
    }
  });
 
});
 
// ── Toggle genérico — llamarlo desde onclick en cualquier vista ──
function togglePanel(panelId, chevronId) {
  var panel   = document.getElementById(panelId);
  var chevron = chevronId ? document.getElementById(chevronId) : null;
  if (!panel) return;
 
  var hidden = panel.classList.contains('hidden');
  panel.classList.toggle('hidden', !hidden);
  if (chevron) chevron.style.transform = hidden ? 'rotate(0deg)' : 'rotate(-90deg)';
}
 