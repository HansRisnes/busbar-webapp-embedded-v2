export const $ = id => document.getElementById(id);

export function ensureFormModal(formId, titleText){
  const form = $(formId);
  if (!form) return null;
  const modalId = `${formId}Modal`;
  let modal = $(modalId);
  if (!modal){
    modal = document.createElement('div');
    modal.id = modalId;
    modal.className = 'form-modal-backdrop';
    modal.hidden = true;
    modal.innerHTML = `
      <div class="form-modal-panel" role="dialog" aria-modal="true" aria-labelledby="${modalId}Title">
        <div class="form-modal-header">
          <h3 id="${modalId}Title"></h3>
          <button type="button" class="form-modal-close" aria-label="Lukk">&#10005;</button>
        </div>
        <div class="form-modal-body"></div>
      </div>
    `;
    document.body.appendChild(modal);
    modal.addEventListener('click', evt=>{
      if (evt.target === modal) closeFormModal(formId);
    });
    modal.querySelector('.form-modal-close')?.addEventListener('click', ()=>closeFormModal(formId));
  }
  const title = modal.querySelector('.form-modal-header h3');
  if (title) title.textContent = titleText || '';
  const body = modal.querySelector('.form-modal-body');
  if (body && form.parentElement !== body){
    body.appendChild(form);
  }
  return modal;
}

export function openFormModal(formId, titleText){
  const modal = ensureFormModal(formId, titleText);
  if (!modal) return;
  modal.hidden = false;
  document.body.classList.add('has-form-modal-open');
}

export function closeFormModal(formId){
  const form = $(formId);
  if (form) form.hidden = true;
  const modal = $(`${formId}Modal`);
  if (modal) modal.hidden = true;
  if (!document.querySelector('.form-modal-backdrop:not([hidden])')){
    document.body.classList.remove('has-form-modal-open');
  }
}
