document.addEventListener('DOMContentLoaded', () => {
  const exportLink = Array.from(document.querySelectorAll('a')).find((link) => {
    const href = link.getAttribute('href') || '';
    return link.textContent.trim() === '下載公版 Excel' && /\/public\.xlsx(?:\?|$)/.test(href);
  });

  if (!exportLink || document.getElementById('public-excel-filename')) return;

  const label = document.createElement('span');
  label.textContent = '檔名';
  label.className = 'muted';

  const input = document.createElement('input');
  input.id = 'public-excel-filename';
  input.type = 'text';
  input.setAttribute('aria-label', '公版 Excel 檔名');
  input.placeholder = '輸入 Excel 檔名';
  input.style.minWidth = '180px';
  input.style.maxWidth = '260px';

  const heading = document.querySelector('.top h1');
  const draftName = heading ? heading.textContent.trim() : '';
  input.value = draftName ? `${draftName}_公版菜單` : '公版菜單';

  exportLink.parentNode.insertBefore(label, exportLink);
  exportLink.parentNode.insertBefore(input, exportLink);

  exportLink.addEventListener('click', () => {
    const filename = input.value.trim();
    const url = new URL(exportLink.href, window.location.origin);
    if (filename) {
      url.searchParams.set('filename', filename);
    } else {
      url.searchParams.delete('filename');
    }
    exportLink.href = url.toString();
  });
});
