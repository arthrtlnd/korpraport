// Format NIK (hanya angka)
document.getElementById('nik').addEventListener('input', function(e) {
    this.value = this.value.replace(/[^0-9]/g, '');
});

// Format No HP (hanya angka)
document.getElementById('no_hp').addEventListener('input', function(e) {
    this.value = this.value.replace(/[^0-9]/g, '');
});

// Logic Matra & Pangkat PNS
const korpByMatra = {
    '1': ['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'H1', 'K1', 'M1', 'N1', 'P1', 'Q1', 'R1', 'X1', 'Y1', 'Z1', 'A3'], 
    '2': ['12', '22', '32', '42', '52', '62', '72', '82'], 
    '3': ['13', '23', '33', '43', '53', '63', '73', '83', '93', 'A3'], 
    '0': [] 
};

const matraSelect = document.getElementById('matra');
const korpSelect = document.getElementById('korp');
const pangkatSelect = document.getElementById('pangkat');
const allKorpOptions = Array.from(korpSelect.options);

function filterKorp() {
    const selectedMatra = matraSelect.value;
    const currentKorp = korpSelect.value;
    
    korpSelect.innerHTML = '<option value="">-- Pilih Korp --</option>';
    
    if (selectedMatra === '0') {
        // Jika PNS, matikan Korp dan Pangkat
        korpSelect.disabled = true;
        korpSelect.value = '';
        if(pangkatSelect) {
            pangkatSelect.disabled = true;
            pangkatSelect.value = '';
        }
    } else if (selectedMatra && korpByMatra[selectedMatra]) {
        korpSelect.disabled = false;
        if(pangkatSelect) pangkatSelect.disabled = false;
        
        allKorpOptions.forEach(option => {
            if (option.value && korpByMatra[selectedMatra].includes(option.value)) {
                const newOption = option.cloneNode(true);
                if (option.value === currentKorp) {
                    newOption.selected = true;
                }
                korpSelect.appendChild(newOption);
            }
        });
    } else {
        korpSelect.disabled = false;
        if(pangkatSelect) pangkatSelect.disabled = false;
        
        allKorpOptions.forEach(option => {
            if (option.value) {
                const newOption = option.cloneNode(true);
                if (option.value === currentKorp) {
                    newOption.selected = true;
                }
                korpSelect.appendChild(newOption);
            }
        });
    }
}

matraSelect.addEventListener('change', filterKorp);
filterKorp();