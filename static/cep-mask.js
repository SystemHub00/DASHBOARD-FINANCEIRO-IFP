// Máscara automática para campo de CEP (formato 00000-000)
document.addEventListener('DOMContentLoaded', function() {
    var cepInput = document.getElementById('cep');
    if (cepInput) {
        cepInput.addEventListener('input', function(e) {
            let v = cepInput.value.replace(/\D/g, '');
            if (v.length > 5) {
                v = v.slice(0, 5) + '-' + v.slice(5, 8);
            }
            cepInput.value = v.slice(0, 9);
        });
    }
});
