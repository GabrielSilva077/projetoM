function saveAsJPEG() {
  const tabela = document.getElementById('tabela');

  // Garantir que a tabela esteja visível e renderizada
  tabela.style.display = 'block';
  tabela.style.margin = '20px auto';  // Adiciona uma margem ao redor da tabela

  html2canvas(tabela).then(function(canvas) {
    // Converte o canvas para uma imagem e força o download
    var link = document.createElement('a');
    link.download = 'igreja.jpeg';
    link.href = canvas.toDataURL('image/jpeg');
    link.click();
    
    // Redefinir o estilo após a captura
    tabela.style.display = '';
    tabela.style.margin = '';
  }).catch(function(error) {
    console.error('Erro ao capturar a imagem:', error);
  });
}

let lastId = 0;

function addRow() {
  const table = document
    .getElementById("excel")
    .getElementsByTagName("tbody")[0];
  const rowId = Number(document.getElementById("addNewRowId").value);
  let newRow;

  const rowToAddBelow = table.querySelector(`tr[data-id="${rowId}"]`);
  if (rowToAddBelow) {
    newRow = rowToAddBelow.cloneNode(true);
  } else {
    newRow = table.lastElementChild.cloneNode(true);
    lastId++;
    newRow.setAttribute("data-id", lastId);
  }

  if (rowToAddBelow) {
    lastId++;
    newRow.setAttribute("data-id", lastId);
    table.insertBefore(newRow, rowToAddBelow.nextSibling);
  } else {
    table.appendChild(newRow);
  }

  updateIds("excel");
  replicateFunctionalities(newRow);
}

function replicateFunctionalities(row) {
  // Clonar funcionalidades para a nova linha aqui
  cloneButtons(row);
  cloneSelects(row);
  // Adicione outras funcionalidades aqui, se necessário
}

function cloneButtons(row) {
  const buttons = row.querySelectorAll("button");
  buttons.forEach((button) => {
    button.onclick = function () {
      button.classList.toggle("selected");
    };
  });
}

function cloneSelects(row) {
  const selects = row.querySelectorAll("select");
  selects.forEach((select) => {
    const options = select.querySelectorAll("option");
    options.forEach((option) => {
      option.onclick = function () {
        option.classList.toggle("selected");
      };
    });
  });
}

function removeRow() {
  const table = document
    .getElementById("excel")
    .getElementsByTagName("tbody")[0];
  const rowId = Number(document.getElementById("removeRowId").value);

  const rowToRemove = table.querySelector(`tr[data-id="${rowId}"]`);
  if (rowToRemove) {
    table.removeChild(rowToRemove);
    updateIds("excel");
  }
}

function fillTable() {
  const startDate = new Date(document.getElementById("startDate").value);
  const endDate = new Date(document.getElementById("endDate").value);
  const table = document
    .getElementById("excel")
    .getElementsByTagName("tbody")[0];
  table.innerHTML = ""; // Limpa o corpo da tabela
  lastId = 0;

  while (startDate <= endDate) {
    const newRow = createNewRow(startDate);
    table.appendChild(newRow);

    const dayOfWeek = startDate.getDay();

    if (dayOfWeek === 0) {
      // Se for domingo, repete a linha 3 vezes
      for (let i = 0; i < 2; i++) {
        const newRepeatedRow = createNewRow(startDate);
        table.appendChild(newRepeatedRow);
      }
    }

    startDate.setDate(startDate.getDate() + 1); // Incrementa a data
  }

  updateIds("excel");
}

function createNewRow(date) {
  const newRow = document.createElement("tr");
  lastId++;
  newRow.setAttribute("data-id", lastId);

  const dayOfWeek = date.getDay();
  const dayName = getDayName(dayOfWeek);

  for (let i = 0; i < 6; i++) {
    const cell = document.createElement("td");
    if (i === 0) {
      cell.textContent = lastId;
    } else if (i === 1) {
      cell.textContent = formatDate(date);
    } else if (i === 2) {
      cell.textContent = dayName;
    } else if (i === 3) {
      const select = document.createElement("select");
      const times = [
        "06:00",
        " ",
        "06:30",
        "07:00",
        "07:30",
        "08:00",
        "08:30",
        "09:00",
        "09:30",
        "10:00",
        "10:30",
        "11:00",
        "11:30",
        "12:00",
        "12:30",
        "13:00",
        "13:30",
        "14:00",
        "14:30",
        "15:00",
        "15:30",
        "16:00",
        "16:30",
        "17:00",
        "17:30",
        "18:00",
        "18:30",
        "19:00",
        "19:30",
        "20:00",
        "20:30",
        "21:00",
        "21:30",
        "22:00",
      ]; // Exemplo de horários
      times.forEach((time) => {
        const option = document.createElement("option");
        option.value = time;
        option.textContent = time;
        select.appendChild(option);
      });
      cell.appendChild(select);
    } else if (i === 4) {
      const names = [
      
        "Maycon",
        "Luan",
        "Diana",
        "Laís",
        "José",
        "Rita",
        "Gisele",
        "Cris",
        "Rodrigo",
        "Erica",
        "Gloria",
        "Cássia",
        "Jennifer",
        "Ana Claudia",
        "Ana Carla",
        "Edesio",
        "Samuel",
        "Sheila",
        "Janaina",
        "Juliana",
        "Keel",
        "Kelly",
        "Mara",
        "Maria Rosa",
        "Miriam",
        "Nathalia",
        "Naná",
        "Nayara",
        "Conceição",
        "Regina",
        "Renata",
        "Rodolfo",
        "Selma",
        "Shyrlei",
        "Sol",        
        "Nilce",
        "Edvanio",
        "Neide",
      ]; // Exemplo de nomes
      names.forEach((name) => {
        const button = document.createElement("button");
        button.textContent = name;
        button.classList.add("person-name"); // Adicionando classe para estilização
        button.onclick = function () {
          button.classList.toggle("selected");
        };
        cell.appendChild(button);
      });
    } else if (i === 5) {
      const missas = [
        "Domingo",
        " ",
        "Jubilar Santa Mãe de Deus",
        "Maria Mãe de Deus",
        "cerco de jericó",
        "Véspera de Natal",
        "Natal",
        "Véspera de Ano Novo",
        "Santa Terezinha",
        "Semana Carismática",
        "Domingo de Pentecostes",
        "Corpus Christi",
        "Semanal",
        "São José",
        "Cura e libertação",
        "Grupo De Oração",
        "Penitencial",
        "SCJ / Santa Terezinha",
        "Aniversário Ordenação Pe.Elvis",
        "Aniversário Ordenação Pe.Eldivar",
        "Santuario Mãe de Deus",
        "Ceia do Senhor",
        "Paixão de Cristo",
        "Vigília Pascoal",
        "Domingo De Páscoa",
        "Triduo Sagrado C. de Jesus",
        "Solenidade Sag. Cor. de Jesus",
        "Novena Sta. Terezinha",
        "Festa de Sta. Terezinha",
        "Chegada da Reliquia",
        "N.Sra Aparecida",
        "N.Sra Aparecida (Campo)",
        "N.Sra Aparecida (Paróquia)",
        "Novena",

      ]; // Tipos de missa
      const selectMissa = document.createElement("select");
      missas.forEach((missa) => {
        const option = document.createElement("option");
        option.value = missa;
        option.textContent = missa;
        selectMissa.appendChild(option);
      });
      cell.appendChild(selectMissa);
    }

    cell.classList.add(getDayClass(dayOfWeek)); // Adiciona classe do dia da semana
    newRow.appendChild(cell);
  }

  return newRow;
}

function formatDate(date) {
  const day = date.getDate();
  return `${(day < 10 ? "0" : "") + day}`;
}

function getDayName(dayOfWeek) {
  const days = [
    "Domingo",
    "Segunda-feira",
    "Terça-feira",
    "Quarta-feira",
    "Quinta-feira",
    "Sexta-feira",
    "Sábado",
  ];
  return days[dayOfWeek];
}

function getDayClass(dayOfWeek) {
  const classes = [
    "domingo",
    "segunda",
    "terca",
    "quarta",
    "quinta",
    "sexta",
    "sabado",
  ];
  return classes[dayOfWeek];
}

function updateIds(tableId) {
  const rows = document.getElementById(tableId).querySelectorAll("tbody tr");
  for (let i = 0; i < rows.length; i++) {
    rows[i].setAttribute("data-id", i + 1);
    rows[i].querySelector("td:first-child").textContent = i + 1;
  }
  lastId = rows.length;
}

function addSelectedNames() {
  const rows = document.querySelectorAll("#excel tbody tr");
  rows.forEach((row) => {
    const buttons = row.querySelectorAll("button.selected");
    const namesCell = row.querySelector("td:nth-child(5)");
    const namesArray = Array.from(buttons).map((button) => button.textContent);
    namesCell.textContent = namesArray.join(", ");
  });
}

function addSelectedMissas() {
  const rows = document.querySelectorAll("#excel tbody tr");
  rows.forEach((row) => {
    const select = row.querySelector("td:nth-child(6) select");
    const missasCell = row.querySelector("td:nth-child(6)");
    const selectedMissas = Array.from(select.options)
      .filter((option) => option.selected)
      .map((option) => option.textContent);
    missasCell.textContent = selectedMissas.join(", ");
  });
}

document.getElementById("excel").addEventListener("click", function (event) {
  const target = event.target;
  if (target.classList.contains("person-name")) {
    const namesTableRows = document
      .getElementById("names-table")
      .querySelectorAll("tbody tr");

    for (let i = 0; i < namesTableRows.length; i++) {
      const nameCell = namesTableRows[i].querySelector("td:first-child");
      const countCell = namesTableRows[i].querySelector("td:nth-child(2)");
      const name = nameCell.textContent;

      if (target.textContent === name) {
        const currentCount = parseInt(countCell.textContent, 10);
        if (target.classList.contains("selected")) {
          countCell.textContent = (
            currentCount + 1 >= 0 ? currentCount + 1 : 0
          ).toString();
        } else {
          countCell.textContent = (currentCount - 1).toString();
        }
        break;
      }
    }
  }
});

// Adicionando a função para criar a tabela de nomes e contadores
function createNamesTable() {
  const names = [
  
    "Maycon",
    "Luan",
    "Diana",
    "Laís",
    "José",
    "Rita",
    "Gisele",
    "Cris",
    "Rodrigo",
    "Erica",
    "Gloria",
    "Cássia",
    "Jennifer",
    "Ana Claudia",
    "Ana Carla",
    "Edesio",
    "Samuel",
    "Sheila",
    "Janaina",
    "Juliana",
    "Keel",
    "Kelly",
    "Mara",
    "Maria Rosa",
    "Miriam",
    "Nathalia",
    "Naná",
    "Nayara",
    "Conceição",
    "Regina",
    "Renata",
    "Rodolfo",
    "Selma",
    "Shyrlei",
    "Sol",
    "Nilce",
    "Edvanio",
    "Neide",
  ];

  const namesTable = document
    .getElementById("names-table")
    .getElementsByTagName("tbody")[0];

  names.forEach((name) => {
    const row = document.createElement("tr");

    const nameCell = document.createElement("td");
    const countCell = document.createElement("td");
    countCell.textContent = "0"; // Inicializando o contador em 0

    nameCell.textContent = name;

    row.appendChild(nameCell);
    row.appendChild(countCell);
    namesTable.appendChild(row);
  });
}

// Chamando a função para criar a tabela de nomes e contadores ao carregar a página
window.addEventListener("load", createNamesTable);

function editNames() {
  const editarNomesBtn = document.getElementById("editarNomesBtn");
  editarNomesBtn.textContent = "Salvar Nomes";
  // Adicione aqui a lógica para editar a tabela conforme sua necessidade

  // Exemplo de lógica: tornar os campos editáveis
  const namesTableRows = document
    .getElementById("names-table")
    .querySelectorAll("tbody tr");
  namesTableRows.forEach((row) => {
    const nameCell = row.querySelector("td:first-child");
    const input = document.createElement("input");
    input.type = "text";
    input.value = nameCell.textContent;
    nameCell.innerHTML = "";
    nameCell.appendChild(input);
  });

  editarNomesBtn.onclick = function () {
    saveEditedNames();
  };
}

function saveEditedNames() {
  const namesTableRows = document
    .getElementById("names-table")
    .querySelectorAll("tbody tr");
  namesTableRows.forEach((row) => {
    const nameCell = row.querySelector("td:first-child");
    const input = nameCell.querySelector("input");
    nameCell.textContent = input.value;
  });

  const editarNomesBtn = document.getElementById("editarNomesBtn");
  editarNomesBtn.textContent = "Editar Tabela";
  editarNomesBtn.onclick = function () {
    editNames();
  };
}