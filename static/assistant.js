/*
Arquivo: static/assistant.js
Descrição: Script para exibir um assistente virtual com instruções passo a passo no site.
*/


// Define mensagens diferentes para cada página
let steps = [];
const path = window.location.pathname;
if (path === '/' || path === '/index' || path === '/index.html') {
  steps = [
    'Olá! Para começar, clique no botão <b>Começar inscrição</b>.',
    'Você vai avançar para o preenchimento dos seus dados. Siga etapa por etapa para concluir.'
  ];
} else if (path.startsWith('/inscricao')) {
  steps = [
    'Preencha seu <b>nome completo</b> exatamente como aparece no seu documento.',
    'Selecione seu <b>Gênero</b> na lista.',
    'Digite seu <b>CPF</b>. O campo será formatado automaticamente.',
    'Informe sua <b>data de nascimento</b>. O campo também será formatado automaticamente.',
    'Preencha seu <b>WhatsApp</b> com DDD para contato.',
    'Digite seu <b>e-mail</b> para receber atualizações importantes.',
    'Quando terminar, clique em <b>Próximo</b> para avançar.'
  ];
} else if (path.startsWith('/endereco')) {
  steps = [
    'Preencha os dados de endereço com atenção para continuar.',
    'Digite o <b>CEP</b> corretamente para facilitar o preenchimento.',
    'Confira as informações antes de seguir para a próxima etapa.'
  ];
} else if (path.startsWith('/curso')) {
  steps = [
    'Agora selecione a opção desejada na etapa atual.',
    'Após a escolha, as informações complementares serão preenchidas automaticamente.',
    'Confira os dados exibidos e clique em <b>Próximo</b> para continuar.'
  ];
} else if (path.startsWith('/revisao')) {
  steps = [
    'Revise todos os dados com atenção antes de finalizar.',
    'Se quiser, preencha o campo <b>Como conheceu</b>.',
    'Antes de concluir, marque a confirmação dos dados e clique em <b>Finalizar inscrição</b>.'
  ];
} else if (path.startsWith('/confirmacao')) {
  steps = [
    'Tudo certo! Guarde o <b>número de protocolo</b> para consulta e acompanhamento.'
  ];
} else {
  steps = [
    'Use o assistente para acompanhar as orientações desta página.'
  ];
}

let currentStep = 0;

function showAssistant() {
  let assistant = document.getElementById('assistant-box');
  if (!assistant) {
    assistant = document.createElement('div');
    assistant.id = 'assistant-box';
    assistant.innerHTML = `
      <div id="assistant-avatar">🤖</div>
      <div id="assistant-text"></div>
      <button id="assistant-prev">Anterior</button>
      <button id="assistant-next">Próximo</button>
      <button id="assistant-close">Fechar</button>
    `;
    document.body.appendChild(assistant);
    document.getElementById('assistant-prev').onclick = prevStep;
    document.getElementById('assistant-next').onclick = nextStep;
    document.getElementById('assistant-close').onclick = closeAssistant;
  }
  updateAssistant();
}


function updateAssistant() {
  // Permite HTML nas mensagens (ex: <b>texto</b>)
  document.getElementById('assistant-text').innerHTML = steps[currentStep];
  document.getElementById('assistant-prev').disabled = currentStep === 0;
  document.getElementById('assistant-next').disabled = currentStep === steps.length - 1;
}

function nextStep() {
  if (currentStep < steps.length - 1) {
    currentStep++;
    updateAssistant();
  }
}

function prevStep() {
  if (currentStep > 0) {
    currentStep--;
    updateAssistant();
  }
}

function closeAssistant() {
  document.getElementById('assistant-box').remove();
}

// Ícone flutuante para abrir o assistente
window.onload = function() {
  // Cria o ícone para abrir o assistente
  let icon = document.createElement('div');
  icon.id = 'assistant-icon';
  icon.innerText = '🤖';
  icon.title = 'Precisa de ajuda? Clique aqui!';
  icon.onclick = showAssistant;
  document.body.appendChild(icon);
};
