const Excel = require('exceljs');

class FileXls {
  static async geraWorksheets({ quebraFolha = 'N', sheets, colunaSheets }) {
    if (quebraFolha === 'N') {
      console.log(quebraFolha);

      return null;
    } else {
      const nomesSheets = sheets.map(s => s[colunaSheets]);

      const result = nomesSheets.filter((nome, valor) => {
        return nomesSheets.indexOf(nome) === valor;
      });

      return result;
    }
  }

  static async gerColunas(colunas) {
    const result = colunas.map(c =>
      Object.create({
        header: c.toLocaleUpperCase(),
        key: c.toLocaleLowerCase(),
        width: 20,
      }),
    );
    return result;
  }

  static async groupBy({ array = [], key = '' }) {
    return array.reduce((acc, item) => {
      if (!acc[item[key]]) acc[item[key]] = [];
      acc[item[key]].push(item);
      return acc;
    }, {});
  }

  static async geraXls({ quebraFolha = 'N', obj }) {
    const workbook = new Excel.Workbook();

    const colunas = await this.gerColunas(Object.keys(obj[0]));

    if (quebraFolha === 'S') {
      const nomesSheaets = await this.geraWorksheets({
        quebraFolha: 'S',
        colunaSheets: 'sexo',
        sheets: obj,
      });

      const arrAgrupado = await this.groupBy({ array: obj, key: 'sexo' });
      const sheets = nomesSheaets.map(s => workbook.addWorksheet(s));

      for (let i = 0; i < sheets.length; i++) {
        sheets[i].columns = colunas;
        sheets[i].addRows(arrAgrupado[sheets[i].name]);
      }

      await workbook.xlsx.writeFile('export.xlsx');
    }

    /*sheets.columns = colunas;
    sheets.addRows(obj);

    */

    return colunas;
  }
}

(async () => {
  const obj = [
    {
      nome: 'joao',
      idade: '16',
      sexo: 'M',
    },
    {
      nome: 'Maria',
      idade: '57',
      sexo: 'F',
    },
    {
      nome: 'Marta',
      idade: '87',
      sexo: 'F',
    },
    {
      nome: 'Pedro',
      idade: '60',
      sexo: 'M',
    },

    {
      nome: 'Alceu',
      idade: '18',
      sexo: 'I',
    },
  ];
  await FileXls.geraXls({ quebraFolha: 'S', obj });
})();
