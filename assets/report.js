/* Shared behaviour for every company page (inlined into <Company>/index.html at build time):
   one tooltip used by all charts, and row + column highlight on tables. */
(function () {
  var tip = document.createElement('div');
  tip.className = 'tip';
  tip.setAttribute('role', 'tooltip');
  document.body.appendChild(tip);

  function place(x, y) {
    var w = tip.offsetWidth, h = tip.offsetHeight;
    var left = x + 14, top = y - h - 12;
    if (left + w > window.innerWidth - 8) left = x - w - 14;
    if (left < 8) left = 8;
    if (top < 8) top = y + 18;
    tip.style.left = left + 'px';
    tip.style.top = top + 'px';
  }

  window.ReportTip = {
    show: function (html, x, y) { tip.innerHTML = html; tip.style.display = 'block'; place(x, y); },
    move: place,
    hide: function () { tip.style.display = 'none'; },
    head: function (t) { return '<div class="th">' + t + '</div>'; },
    row: function (name, value, color, cls) {
      return '<div' + (cls ? ' class="' + cls + '"' : '') + '>' +
        (color ? '<i style="background:' + color + '"></i>' : '') + name + '<b>' + value + '</b></div>';
    }
  };

  document.querySelectorAll('.tw table').forEach(function (t) {
    t.addEventListener('mouseover', function (ev) {
      var td = ev.target.closest('td');
      t.querySelectorAll('td.xcol').forEach(function (c) { c.classList.remove('xcol'); });
      if (!td || td.cellIndex === 0 || td.parentElement.classList.contains('sub')) return;
      t.querySelectorAll('tbody tr:not(.sub)').forEach(function (tr) {
        if (tr.cells[td.cellIndex]) tr.cells[td.cellIndex].classList.add('xcol');
      });
    });
    t.addEventListener('mouseleave', function () {
      t.querySelectorAll('td.xcol').forEach(function (c) { c.classList.remove('xcol'); });
    });
  });

  window.addEventListener('scroll', function () { tip.style.display = 'none'; }, { passive: true });
})();
