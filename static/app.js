document.querySelectorAll('[data-file-input]').forEach(input => {
  const filename = input.parentElement.querySelector('[data-file-name]');
  input.addEventListener('change', () => {
    filename.textContent = input.files && input.files.length ? input.files[0].name : '未选择文件';
  });
});
document.querySelectorAll('form[data-confirm]').forEach(form => form.addEventListener('submit', event => {
  if (!window.confirm(form.dataset.confirm)) event.preventDefault();
}));
document.querySelectorAll('form[data-loading]').forEach(form => form.addEventListener('submit', event => {
  if (event.defaultPrevented) return;
  const button = form.querySelector('button[type="submit"],button:not([type])');
  if (button) {
    button.disabled = true;
    button.textContent = '处理中…';
  }
}));
