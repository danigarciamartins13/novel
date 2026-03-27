/* ── Máscara de telefone ─────────────────────────────────────────────────── */
function aplicarMascaraTel(input) {
  input.addEventListener('input', function () {
    let v = this.value.replace(/\D/g, '').slice(0, 11);
    if (v.length > 10)      v = v.replace(/^(\d{2})(\d{5})(\d{4})$/, '($1) $2-$3');
    else if (v.length > 6)  v = v.replace(/^(\d{2})(\d{4})(\d{0,4})$/, '($1) $2-$3');
    else if (v.length > 2)  v = v.replace(/^(\d{2})(\d{0,5})$/, '($1) $2');
    else if (v.length > 0)  v = '(' + v;
    this.value = v;
  });
}

// Aplica no campo fixo
const telFixo = document.getElementById('telefone');
if (telFixo) aplicarMascaraTel(telFixo);

// Aplica em todos os campos extras do tipo tel
document.querySelectorAll('.campo-tel').forEach(aplicarMascaraTel);

/* ── Submit do formulário ────────────────────────────────────────────────── */
document.getElementById('formInscricao').addEventListener('submit', async function (e) {
  e.preventDefault();
  const btn = document.getElementById('btnSubmit');
  const msg = document.getElementById('formMsg');
  btn.disabled = true;
  btn.textContent = 'Enviando…';
  msg.textContent = '';
  msg.className = 'form-msg';

  // Coleta campos padrão
  const payload = {
    nome:     document.getElementById('nome').value.trim(),
    email:    document.getElementById('email').value.trim(),
    telefone: document.getElementById('telefone').value.trim(),
    extras:   {},
  };

  // Coleta campos extras pelo atributo data-campo-id
  document.querySelectorAll('[data-campo-id]').forEach(function (el) {
    payload.extras[el.dataset.campoId] = el.value.trim();
  });

  try {
    const res  = await fetch('/inscrever', {
      method:  'POST',
      headers: { 'Content-Type': 'application/json' },
      body:    JSON.stringify(payload),
    });
    const data = await res.json();

    if (data.ok) {
      document.getElementById('formInscricao').reset();
      abrirModal();
    } else {
      msg.textContent = data.erro || 'Ocorreu um erro. Tente novamente.';
      msg.className = 'form-msg erro';
    }
  } catch (_) {
    msg.textContent = 'Falha na conexão. Tente novamente.';
    msg.className = 'form-msg erro';
  } finally {
    btn.disabled = false;
    // Restaura o texto original do botão
    const btnOrig = btn.getAttribute('data-texto') || 'Quero me inscrever';
    btn.textContent = btnOrig;
  }
});

// Guarda o texto original do botão para restaurar após submit
(function () {
  const btn = document.getElementById('btnSubmit');
  if (btn) btn.setAttribute('data-texto', btn.textContent);
})();

/* ── Modal ───────────────────────────────────────────────────────────────── */
function abrirModal()  { document.getElementById('modalSucesso').classList.add('show'); }
function fecharModal() { document.getElementById('modalSucesso').classList.remove('show'); }

document.getElementById('modalSucesso').addEventListener('click', function (e) {
  if (e.target === this) fecharModal();
});
