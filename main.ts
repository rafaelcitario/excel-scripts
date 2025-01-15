import { exec } from 'child_process';
import fs from 'fs';
import xlsx from 'node-xlsx';

export interface IAccountObj {
  account: {
    id: string;
    data: { [x: string]: Object };
  };
}


export function jsonfyDataExcel() {
  const workSheetFromFile = xlsx.parse('./src/files/Histórico dos registros do objeto Conta.xlsx');
  const account: IAccountObj[] = [];
  
  for (let i = 1; i < workSheetFromFile[0].data.length; i++) { 
    const line = workSheetFromFile[0].data[i];
    const accountObj: IAccountObj = {
      account: {
        id: line[1],
        data: {},
      },
    };
    
    for (let j = 2; j < workSheetFromFile[0].data[0].length; j++) {
      const keyColl = workSheetFromFile[0].data[0][j];
      accountObj.account.data[keyColl] = line[j];
    }
    
    account.push(accountObj);
  }
  
  return account;
}


function checkAndCreateDirectory(callback: Function) {
  const directoryPath = './src/jsonData';

  fs.access(directoryPath, (err) => {
    if (!err) {
      
      console.log('Diretório já existe!');
      callback();
    } else {
      
      const command = process.platform === 'win32' 
        ? `mkdir "${directoryPath}"`  
        : `mkdir -p ${directoryPath}`;  

      exec(command, (err, stdout, stderr) => {
        if (err) {
          console.error(`Erro ao criar diretório: ${stderr}`);
          return;
        }
        console.log('Diretório criado com sucesso!');
        callback();
      });
    }
  });
}


function createJsonDataExcel() {
  
  checkAndCreateDirectory(() => {
    const jsonData = jsonfyDataExcel();

    jsonData.forEach((data) => {
      const filePath = `./src/jsonData/${data.account.id}.json`;

      
      fs.writeFile(filePath, JSON.stringify(data), 'utf-8', (err) => {
        if (err) {
          console.error(`Erro ao salvar o arquivo ${filePath}:`, err);
          return;
        }
        console.log(`Arquivo ${filePath} salvo com sucesso!`);
      });
    });
  });
}


createJsonDataExcel();
