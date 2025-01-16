import { exec } from 'child_process';
import fs from 'fs';
import xlsx from 'node-xlsx';

export interface IAccountObj {
  id: string;
  data: { [x: string]: Object; };
};



export function jsonfyDataExcel ( arquivoParaConversao: string, numeroDeArquivosConvertidos: number = 0 ) {
  const workSheetFromFile = xlsx.parse( `./src/plurix/excelFiles/${arquivoParaConversao}` );
  const account: IAccountObj[] = [];

  numeroDeArquivosConvertidos <= 0 ? workSheetFromFile[0].data.length : numeroDeArquivosConvertidos++;

  for ( let i = 1; i < numeroDeArquivosConvertidos; i++ ) {
    const line = workSheetFromFile[0].data[i];
    const accountObj: IAccountObj = {
      id: line[1],
      data: {},
    };

    for ( let j = 0; j < workSheetFromFile[0].data[0].length; j++ ) {
      if ( j == 1 ) continue;
      const keyColl = workSheetFromFile[0].data[0][j];
      accountObj.data[keyColl] = line[j];
    }

    account.push( accountObj );
  }

  return account;
}


function checkAndCreateDirectory ( nomeDaPasta: string, callback: Function ) {
  const directoryPath = `./src/plurix/files/${nomeDaPasta}`;

  fs.access( directoryPath, ( err ) => {
    if ( !err ) {

      console.log( 'Diretório já existe!' );
      callback();
    } else {

      const command = process.platform === 'win32'
        ? `mkdir "${directoryPath}"`
        : `mkdir -p ${directoryPath}`;

      exec( command, ( err, stdout, stderr ) => {
        if ( err ) {
          console.error( `Erro ao criar diretório: ${stderr}` );
          return;
        }
        console.log( 'Diretório criado com sucesso!' );
        callback();
      } );
    }
  } );
}


function createJsonDataExcel ( arquivoParaConversao: string, numeroDeArquivosConvertidos: number = 0, nomeDaPasta: string ) {

  checkAndCreateDirectory( nomeDaPasta, () => {
    const jsonData = jsonfyDataExcel( arquivoParaConversao, numeroDeArquivosConvertidos );

    jsonData.forEach( ( data ) => {
      const filePath = `./src/plurix/files/${nomeDaPasta}/${data.id}.json`;


      fs.writeFile( filePath, JSON.stringify( data ), 'utf-8', ( err ) => {
        if ( err ) {
          console.error( `Erro ao salvar o arquivo ${filePath}:`, err );
          return;
        }
        console.log( `Arquivo ${filePath} salvo com sucesso!` );
      } );
    } );
  } );
}


createJsonDataExcel( "Histórico de registros do objeto caso.xlsx", 3, "cases" );
