import { exec } from 'child_process';
import fs from 'fs/promises';
import xlsx from 'node-xlsx';
import pLimit from 'p-limit';

export interface IAccountObj {
  id: string;
  data: { [x: string]: Object; };
};

export async function jsonfyDataExcel ( arquivoParaConversao: string, numeroDeArquivosConvertidos: number = 0 ) {
  const workSheetFromFile = xlsx.parse( `./src/plurix/excelFiles/${arquivoParaConversao}` );
  const account: IAccountObj[] = [];

  const numFiles = numeroDeArquivosConvertidos <= 0 ? workSheetFromFile[0].data.length : numeroDeArquivosConvertidos + 1;

  for ( let i = 1; i < numFiles; i++ ) {
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


async function checkAndCreateDirectory ( nomeDaPasta: string ): Promise<void> {
  const directoryPath = `./src/plurix/files/${nomeDaPasta}`;

  try {
    await fs.access( directoryPath );
    console.log( 'Diretório já existe!' );
  } catch ( err ) {
    const command = process.platform === 'win32'
      ? `mkdir "${directoryPath}"`
      : `mkdir -p ${directoryPath}`;

    await new Promise<void>( ( resolve, reject ) => {
      exec( command, ( err, stdout, stderr ) => {
        if ( err ) {
          reject( `Erro ao criar diretório: ${stderr}` );
        } else {
          console.log( 'Diretório criado com sucesso!' );
          resolve();
        }
      } );
    } );
  }
}


async function createJsonDataExcel ( arquivoParaConversao: string, nomeDaPasta: string, numeroDeArquivosConvertidos: number = 0 ) {
  try {
    await checkAndCreateDirectory( nomeDaPasta );
    const jsonData = await jsonfyDataExcel( arquivoParaConversao, numeroDeArquivosConvertidos );


    const limit = pLimit( 100 );


    const writePromises = jsonData.map( ( data ) =>
      limit( async () => {
        const filePath = `./src/plurix/files/${nomeDaPasta}/${data.id}.json`;
        await fs.writeFile( filePath, JSON.stringify( data ), 'utf-8' );
        console.log( `Arquivo ${filePath} salvo com sucesso!` );
      } )
    );


    await Promise.all( writePromises );

  } catch ( error ) {
    console.error( 'Erro ao processar os arquivos:', error );
  }
}


createJsonDataExcel( "Histórico de registros do objeto caso.xlsx", "cases", 0 );
createJsonDataExcel( "Histórico dos registros do objeto Conta.xlsx", "accounts", 0 );
createJsonDataExcel( "Histórico dos registros do objeto contato.xlsx", "contacts", 0 );
createJsonDataExcel( "Histórico dos registros do objeto histórico do caso.xlsx", "caseHistory", 0 );
createJsonDataExcel( "Histórico dos registros do objeto transcrições de chat.xlsx", "transcription", 0 );
