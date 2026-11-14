/* Shared behaviour for every company page (inlined into <Company>/index.html at build time):
   one tooltip used by all charts, row + column highlight on tables, and Excel-style cell selection. */
(function () {
  var sel = null;                                      // the selected table cell, or null when nothing is selected
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

  // true when the reader uses a finger, not a mouse: charts then open the tooltip on tap
  var touch = window.matchMedia('(pointer: coarse)').matches;

  window.ReportTip = {
    touch: touch,
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
      if (sel) return;                                 // a selected cell is the only highlight, as in Excel
      if (!td || td.cellIndex === 0 || td.parentElement.classList.contains('sub')) return;
      t.querySelectorAll('tbody tr:not(.sub)').forEach(function (tr) {
        if (tr.cells[td.cellIndex]) tr.cells[td.cellIndex].classList.add('xcol');
      });
    });
    t.addEventListener('mouseleave', function () {
      t.querySelectorAll('td.xcol').forEach(function (c) { c.classList.remove('xcol'); });
    });
  });

  // ---- Excel-style selection: click a cell, then move with the arrow keys ----
  // The selected cell gets a solid outline, its row and column stay highlighted (the cross), and the
  // table scrolls itself so the cell stays in view: no dragging the scrollbar at the bottom of a long table.

  function dataRows(t) {                               // the rows that hold figures (STEP rows are skipped)
    return Array.prototype.filter.call(t.tBodies[0].rows, function (r) { return !r.classList.contains('sub'); });
  }

  function clearSel() {                                // remove the outline and the cross
    if (!sel) return;
    sel.closest('table').querySelectorAll('td.scol,td.srow,td.sel').forEach(function (c) {
      c.classList.remove('scol', 'srow', 'sel');
    });
    sel = null;
    document.body.classList.remove('has-sel');        // hovering highlights rows again
  }

  function reveal(td) {                                // scroll the table box, then the page, so the cell shows
    var box = td.closest('.tw');
    var pin = td.parentElement.cells[0].offsetWidth;   // the pinned label column covers the left edge
    var left = td.offsetLeft, right = left + td.offsetWidth;
    if (left - pin < box.scrollLeft) box.scrollLeft = left - pin;
    else if (right > box.scrollLeft + box.clientWidth) box.scrollLeft = right - box.clientWidth;
    var r = td.getBoundingClientRect();
    if (r.top < 70) window.scrollBy(0, r.top - 70);
    else if (r.bottom > window.innerHeight - 30) window.scrollBy(0, r.bottom - window.innerHeight + 30);
  }

  function select(td) {                                // outline one cell, highlight its row and column
    clearSel();
    sel = td;
    document.body.classList.add('has-sel');           // hovering stops highlighting while a cell is selected
    document.querySelectorAll('td.xcol').forEach(function (c) { c.classList.remove('xcol'); });
    var i = td.cellIndex;
    dataRows(td.closest('table')).forEach(function (tr) { if (tr.cells[i]) tr.cells[i].classList.add('scol'); });
    Array.prototype.forEach.call(td.parentElement.cells, function (c) { c.classList.add('srow'); });
    td.classList.add('sel');
    reveal(td);
  }

  document.querySelectorAll('.tw table').forEach(function (t) {
    t.addEventListener('click', function (ev) {
      var td = ev.target.closest('td');
      if (!td || td.parentElement.classList.contains('sub')) return;   // STEP rows are titles, not cells
      if (td.cellIndex === 0) td = td.parentElement.cells[1];         // a click on a label picks its first figure
      if (td) select(td);
    });
  });

  document.addEventListener('keydown', function (ev) {
    if (!sel) return;
    if (ev.key === 'Escape') { clearSel(); return; }
    var tr = sel.parentElement, i = sel.cellIndex, next = null;
    var rows = dataRows(sel.closest('table')), r = rows.indexOf(tr);
    if (ev.key === 'ArrowLeft' && i > 1) next = tr.cells[i - 1];
    else if (ev.key === 'ArrowRight' && tr.cells[i + 1] && !tr.cells[i + 1].classList.contains('fx'))
      next = tr.cells[i + 1];                          // the hidden Formula column is never selected
    else if (ev.key === 'ArrowUp' && r > 0) next = rows[r - 1].cells[i];
    else if (ev.key === 'ArrowDown' && r < rows.length - 1) next = rows[r + 1].cells[i];
    if (ev.key.indexOf('Arrow') === 0) ev.preventDefault();             // the arrows move the cell, not the page
    if (next) select(next);
  });

  // a click or tap outside every table lets the selection go
  document.addEventListener('pointerdown', function (ev) {
    if (sel && !ev.target.closest('.tw table')) clearSel();
  }, { passive: true });

  window.addEventListener('scroll', function () { tip.style.display = 'none'; }, { passive: true });

  // a tap anywhere outside a chart closes the tooltip, so it never stays stuck on a phone
  document.addEventListener('pointerdown', function (ev) {
    if (!ev.target.closest('svg[data-chart]')) tip.style.display = 'none';
  }, { passive: true });

  // on touch screens the pages say "Tap", not "Hover"
  if (touch) document.querySelectorAll('.hovertap').forEach(function (el) { el.textContent = 'Tap'; });
})();
