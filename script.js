document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('fileInput');
    const loadButton = document.getElementById('loadButton');

    loadButton.addEventListener('click', () => {
        const file = fileInput.files[0];
        if (!file) { alert('Por favor, selecione um arquivo.'); return; }

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                processData(jsonData);
            } catch {
                alert('Não foi possível ler o arquivo. Certifique-se de que é um arquivo Excel válido.');
            }
        };
        reader.readAsArrayBuffer(file);
    });
});

function processData(data) {
    if (!Array.isArray(data) || data.length < 2) {
        alert('O arquivo não contém dados suficientes.');
        return;
    }

    const header = data[0];
    const nomeIndex = header.indexOf('Nome completo');
    const dataNascIndex = header.indexOf('Data de nascimento');

    if (nomeIndex === -1 || dataNascIndex === -1) {
        alert('Os cabeçalhos "Nome completo" e "Data de nascimento" não foram encontrados.');
        return;
    }

    const today = new Date();
    const todayBirthdays = [];
    const monthBirthdays = [];

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row || row.length === 0) continue;

        const person = {};
        header.forEach((colName, index) => {
            person[colName] = row[index];
        });

        const nome = person['Nome completo'];
        let dataNascimento = person['Data de nascimento'];

        if (!nome || !dataNascimento) continue;

        const birthDate = parseBirthDate(dataNascimento);
        if (!birthDate) continue;

        person['Data de nascimento'] = formatDate(birthDate);

        if (birthDate.getDate() === today.getDate() && birthDate.getMonth() === today.getMonth()) {
            todayBirthdays.push(person);
        }

        if (birthDate.getMonth() === today.getMonth()) {
            monthBirthdays.push(person);
        }
    }

    displayBirthdays('todayBirthdays', todayBirthdays);
    displayBirthdays('monthBirthdays', monthBirthdays);
}

function parseBirthDate(value) {
    let birthDate;

    if (value instanceof Date) {
        birthDate = value;
    } else if (typeof value === 'string') {
        const [day, month, year] = value.split('/').map(Number);
        if (!day || !month || !year) return null;
        birthDate = new Date(year, month - 1, day);
    } else if (typeof value === 'number') {
        const date = XLSX.SSF.parse_date_code(value);
        if (!date) return null;
        birthDate = new Date(date.y, date.m - 1, date.d);
    } else {
        return null;
    }

    return isNaN(birthDate.getTime()) ? null : birthDate;
}

function formatDate(date) {
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    return `${day}/${month}/${year}`;
}

function displayBirthdays(elementId, birthdays) {
    const list = document.getElementById(elementId);
    list.innerHTML = '';

    if (birthdays.length > 0) {
        birthdays.forEach(person => {
            const li = document.createElement('li');
            li.textContent = Object.entries(person)
                .filter(([_, value]) => value != null && value !== '')
                .map(([key, value]) => `${key}: ${value}`)
                .join(' | ');
            list.appendChild(li);
        });
    } else {
        list.innerHTML = '<li>Sem aniversariantes.</li>';
    }
}
