/**
 * desplegable_destinos.js
 * Visualizaciones para Seguro Viajero:
 * 1. Gráfico de barras Top 10 Destinos
 * 2. Cards TOP 5 destinos
 * 3. Timeline mensual de solicitudes
 * 4. Mapa de calor por mes/año
 * Requiere: Chart.js
 */

document.addEventListener('DOMContentLoaded', function () {

  // ── Desplegable ──────────────────────────────────────
  const panel   = document.getElementById('destinos-panel');
  const chevron = document.getElementById('destinos-chevron');
  const header  = document.getElementById('destinos-header');

  if (header && panel && chevron) {
    header.addEventListener('click', function () {
      const oculto = panel.style.display === 'none';
      panel.style.display = oculto ? 'block' : 'none';
      chevron.style.transform = oculto ? 'rotate(0deg)' : 'rotate(-90deg)';
    });
  }

  // ── 1. Gráfico Top 10 Destinos ───────────────────────
  const canvasDestinos = document.getElementById('chartDestinos');
  if (canvasDestinos && window.chartDestinosData) {
    const d = window.chartDestinosData;
    new Chart(canvasDestinos.getContext('2d'), {
      type: 'bar',
      data: {
        labels: d.labels,
        datasets: [{
          label: 'Solicitudes',
          data: d.values,
          backgroundColor: 'rgba(0,63,183,0.85)',
          borderColor: '#003FB7',
          borderWidth: 1,
          borderRadius: 6,
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        indexAxis: 'y',
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: ctx => ` ${ctx.raw} solicitudes` } }
        },
        scales: {
          x: { beginAtZero: true, grid: { color: 'rgba(0,0,0,0.05)' }, ticks: { font: { size: 11 }, color: '#64748b' } },
          y: { grid: { display: false }, ticks: { font: { size: 12 }, color: '#1e293b' } }
        }
      }
    });
  }

  // ── 2. Cards Top 5 destinos ──────────────────────────
  const cardsEl = document.getElementById('cardsTopDestinos');
  if (cardsEl && window.chartDestinosData) {
    const d     = window.chartDestinosData;
    const top5  = d.labels.slice(0, 5);
    const vals  = d.values.slice(0, 5);
    const total = vals.reduce((a, b) => a + b, 0);
    const emojis = ['🥇','🥈','🥉','4️⃣','5️⃣'];
    const colors = ['#003FB7','#1d4ed8','#2563eb','#3b82f6','#60a5fa'];

    let html = '<div class="grid grid-cols-2 sm:grid-cols-5 gap-3">';
    top5.forEach((label, i) => {
      const pct = total ? Math.round((vals[i] / total) * 100) : 0;
      html += `
        <div class="rounded-2xl border border-outline-variant/20 p-4 flex flex-col gap-2 text-center"
             style="border-top:3px solid ${colors[i]}">
          <span class="text-2xl">${emojis[i]}</span>
          <p class="text-xs font-bold text-on-surface leading-tight">${label}</p>
          <p class="text-2xl font-bold" style="color:${colors[i]}">${vals[i]}</p>
          <p class="text-[10px] text-slate-400">${pct}% del top 10</p>
        </div>`;
    });
    html += '</div>';
    cardsEl.innerHTML = html;
  }

  // ── 3. Timeline mensual ──────────────────────────────
  const canvasTimeline = document.getElementById('chartTimeline');
  if (canvasTimeline && window.chartTimelineData) {
    const t = window.chartTimelineData;
    new Chart(canvasTimeline.getContext('2d'), {
      type: 'line',
      data: {
        labels: t.labels,
        datasets: [{
          label: 'Solicitudes',
          data: t.values,
          borderColor: '#003FB7',
          backgroundColor: 'rgba(0,63,183,0.08)',
          borderWidth: 2.5,
          pointBackgroundColor: '#003FB7',
          pointBorderColor: '#ffffff',
          pointBorderWidth: 2,
          pointRadius: 4,
          pointHoverRadius: 7,
          fill: true,
          tension: 0.4,
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: ctx => ` ${ctx.raw} solicitudes`,
              title: ctx => ctx[0].label
            }
          }
        },
        scales: {
          x: {
            grid: { color: 'rgba(0,0,0,0.04)' },
            ticks: { font: { size: 10 }, color: '#64748b', maxRotation: 45 }
          },
          y: {
            beginAtZero: true,
            grid: { color: 'rgba(0,0,0,0.04)' },
            ticks: { font: { size: 11 }, color: '#64748b', stepSize: 5 }
          }
        }
      }
    });
  }

  // ── 4. Mapa de calor ─────────────────────────────────
  const heatmapEl = document.getElementById('heatmapDestinos');
  if (heatmapEl && window.chartHeatmapData) {
    const h     = window.chartHeatmapData;
    const meses = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
    const allVals = h.anios.flatMap(a => Object.values(h.data[a] || {}));
    const maxVal  = Math.max(...allVals, 1);

    function cellColor(val) {
      if (!val) return { bg: '#F8FAFC', text: '#CBD5E1' };
      const ratio = val / maxVal;
      const r = Math.round(0   + (1 - ratio) * 219);
      const g = Math.round(63  + (1 - ratio) * 156);
      const b = Math.round(183 + (1 - ratio) * 55);
      return {
        bg: `rgb(${r},${g},${b})`,
        text: ratio > 0.45 ? '#ffffff' : '#1e293b'
      };
    }

    let html = '<div class="overflow-x-auto"><table class="w-full text-xs border-collapse" style="border-spacing:3px">';
    html += '<thead><tr><th class="pb-2 pr-3 text-right text-slate-400 font-semibold" style="width:48px"></th>';
    meses.forEach(m => {
      html += `<th class="pb-2 text-center text-slate-400 font-semibold" style="min-width:36px">${m}</th>`;
    });
    html += '</tr></thead><tbody>';

    h.anios.forEach(anio => {
      html += `<tr><td class="pr-3 py-1 text-right text-slate-600 font-bold text-xs">${anio}</td>`;
      for (let mes = 1; mes <= 12; mes++) {
        const val   = (h.data[anio] && h.data[anio][mes]) || 0;
        const color = cellColor(val);
        html += `
          <td class="py-1 px-0.5">
            <div title="${val ? val + ' solicitudes' : 'Sin datos'}"
                 style="background:${color.bg};color:${color.text};border-radius:6px;
                        padding:5px 2px;text-align:center;font-weight:${val?'600':'400'};
                        cursor:default;transition:opacity .15s"
                 onmouseover="this.style.opacity='.75'"
                 onmouseout="this.style.opacity='1'">
              ${val || '·'}
            </div>
          </td>`;
      }
      html += '</tr>';
    });

    html += '</tbody></table></div>';

    // Leyenda
    html += `
      <div class="flex items-center gap-2 mt-3 justify-end">
        <span class="text-[10px] text-slate-400">Menos</span>
        ${[0, 0.2, 0.4, 0.6, 0.8, 1].map(r => {
          const rv = Math.round(0  + (1-r)*219);
          const gv = Math.round(63 + (1-r)*156);
          const bv = Math.round(183+(1-r)*55);
          return `<div style="width:18px;height:18px;border-radius:4px;background:rgb(${rv},${gv},${bv})"></div>`;
        }).join('')}
        <span class="text-[10px] text-slate-400">Más</span>
      </div>`;

    heatmapEl.innerHTML = html;
  }

});