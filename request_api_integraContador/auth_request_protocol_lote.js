const axios = require('axios');
const fs = require('fs');
const path = require('path');

// Configurações
const API_URL = 'https://gateway.apiserpro.serpro.gov.br/integra-contador/v1/Apoiar';
const ACCESS_TOKEN = 'eyJ4NXQiOiJNalJqWkRRMU1EQmtPR1JqWW1Jek9EVmxaRFEzWkdFeU1EVTVabU5rWldVeU9XUmhPRFZpTnciLCJraWQiOiJNRGN5WVdFNU16ZzVaVFJrT1dVME1XWTBNVE0xTkdJMllqbG1OVFZrT0dJd01UVmlORGRpWldJM1pEUmpZVEpsTTJaa05Ua3dNR1F3TWpZeVlXTmxNQV9SUzI1NiIsImFsZyI6IlJTMjU2In0.eyJzdWIiOiJhdXRlbnRpa3VzIiwiYXV0IjoiQVBQTElDQVRJT04iLCJhdWQiOiJPVEpqd1dscnJSNFJmRENDNXBKMTV4X09KaGNhIiwibmJmIjoxNzMzNzY0Njk4LCJhenAiOiJPVEpqd1dscnJSNFJmRENDNXBKMTV4X09KaGNhIiwic2NvcGUiOiJkZWZhdWx0IiwiaXNzIjoiaHR0cHM6XC9cL3B1Ymxpc2hlci5hcGlzZXJwcm8uc2VycHJvLmdvdi5icjo0NDNcL29hdXRoMlwvdG9rZW4iLCJyZWFsbSI6eyJzaWduaW5nX3RlbmFudCI6ImNhcmJvbi5zdXBlciJ9LCJleHAiOjE3MzM3NjgyOTgsImlhdCI6MTczMzc2NDY5OCwianRpIjoiOWJlYzBhNmItODdkYS00MzRkLTk4MTUtMjg5NjE4ZjMwZjFmIn0.nPOqv5dX8tHaTVTyd4SNOWE9HvmYsLKsp7MP-u1tif98j4qw9kqqhxpDBlsrTt3d4AR_QH-qHxg7bNznNyzUM2QW4y4oc7L-usp0uaJYQgM-9dr_JywuTIZ3pC3MFN2nJoKgpu0FyFoesuHrpPIReIGIHtQyeOMSmnDSSYnQlqiTbdOjol51xRE_NDTDcaHmk9sBYPUzYqgVl6Em0yoPfKl-fURvDZkWfQk4pLAYQTtnr6u_8Hmfj0R_GTlOJHi-eZYkWIR_xNVbHiV5gfrKG0XkwdWP1JkY6AkeuV_lVbt2JykS78C3HDQhu0Tp2caulXPwytSrpBdl0gB2Ul1brw'; // Substitua pelo token de acesso real
const JWT_TOKEN = 'eyJhbGciOiJSUzI1NiJ9.eyJzdWIiOiIwNjMxMDcxMTAwMDE0OSIsInRpcG8iOiJQSiIsImFjY2VzcyI6ImV5SjROWFFpT2lKTmFsSnFXa1JSTVUxRVFtdFBSMUpxV1cxSmVrOUVWbXhhUkZFeldrZEZlVTFFVlRWYWJVNXJXbGRWZVU5WFVtaFBSRlpwVG5jaUxDSnJhV1FpT2lKTlJHTjVXVmRGTlUxNlp6VmFWRkpyVDFkVk1FMVhXVEJOVkUweFRrZEpNbGxxYkcxT1ZGWnJUMGRKZDAxVVZtbE9SR1JwV2xkSk0xcEVVbXBaVkVwc1RUSmFhMDVVYTNkTlIxRjNUV3BaZVZsWFRteE5RVjlTVXpJMU5pSXNJbUZzWnlJNklsSlRNalUySW4wLmV5SnpkV0lpT2lKaGRYUmxiblJwYTNWeklpd2lZWFYwSWpvaVFWQlFURWxEUVZSSlQwNGlMQ0poZFdRaU9pSlBWRXBxZDFkc2NuSlNORkptUkVORE5YQktNVFY0WDA5S2FHTmhJaXdpYm1KbUlqb3hOek16TnpZME5qazRMQ0poZW5BaU9pSlBWRXBxZDFkc2NuSlNORkptUkVORE5YQktNVFY0WDA5S2FHTmhJaXdpYzJOdmNHVWlPaUprWldaaGRXeDBJaXdpYVhOeklqb2lhSFIwY0hNNlhDOWNMM0IxWW14cGMyaGxjaTVoY0dselpYSndjbTh1YzJWeWNISnZMbWR2ZGk1aWNqbzBORE5jTDI5aGRYUm9NbHd2ZEc5clpXNGlMQ0p5WldGc2JTSTZleUp6YVdkdWFXNW5YM1JsYm1GdWRDSTZJbU5oY21KdmJpNXpkWEJsY2lKOUxDSmxlSEFpT2pFM016TTNOamd5T1Rnc0ltbGhkQ0k2TVRjek16YzJORFk1T0N3aWFuUnBJam9pT1dKbFl6QmhObUl0T0Rka1lTMDBNelJrTFRrNE1UVXRNamc1TmpFNFpqTXdaakZtSW4wLm5QT3F2NWRYOHRIYVRWVHlkNFNOT1dFOUh2bVlzTEtzcDdNUC11MXRpZjk4ajRxdzlrcXFoeHBEQmxzclR0M2Q0QVJfUUgtcUh4ZzdiTnpuTnl6VU0yUVc0eTRvYzdMLXVzcDB1YUpZUWdNLTlkcl9KeXd1VElaM3BDM01GTjJuSm9LZ3B1MEZ5Rm9lc3VIcnBQSVJlSUdJSHRReWVPTVNtbkRTU1luUWxxaVRiZE9qb2w1MXhSRV9ORFREY2FIbWs5c0JZUFV6WXFnVmw2RW0weW9QZktsLWZVUnZEWmtXZlFrNHBMQVlRVHRucjZ1XzhIbWZqMFJfR1RsT0pIaS1lWllrV0lSX3hOVmJIaVY1Z2ZyS0cwWGt3ZFdQMUprWTZBa2V1Vl9sVmJ0Mkp5a1M3OEMzSERRaHUwVHAyY2F1bFhQd3l0U3JwQmRsMGdCMlVsMWJydyIsImNlcnRpZmljYXRlRGV0YWlscyI6eyJzZXJpYWxOdW1iZXIiOiI4MDg2ODUyNTU0MjI1OTg2MjM3IiwiaXNzdWVyRE4iOiJDTj1BQyBTQUZFV0VCIFJGQiB2NSwgT1U9U2VjcmV0YXJpYSBkYSBSZWNlaXRhIEZlZGVyYWwgZG8gQnJhc2lsIC0gUkZCLCBPPUlDUC1CcmFzaWwsIEM9QlIiLCJ2YWxpZFVudGlsIjoiMjAyNS0xMC0wMVQxNTowMToxMS0wMzAwIiwidmFsaWRGcm9tIjoiMjAyNC0xMC0wMVQxNTowMToxMS0wMzAwIiwidmVyc2lvbiI6IjMiLCJzdWJqZWN0RE4iOiJDTj1DQU5FTExBIEUgU0FOVE9TIENPTlRBQklMSURBREUgTFREQTowNjMxMDcxMTAwMDE0OSwgT1U9dmlkZW9jb25mZXJlbmNpYSwgT1U9NDg5MTU2NTAwMDAxOTMsIE9VPVJGQiBlLUNOUEogQTEsIE9VPVNlY3JldGFyaWEgZGEgUmVjZWl0YSBGZWRlcmFsIGRvIEJyYXNpbCAtIFJGQiwgTD1WT0xUQSBSRURPTkRBLCBTVD1SSiwgTz1JQ1AtQnJhc2lsLCBDPUJSIn0sImlwIjoiMTc3LjEzMS4xODUuMTAiLCJyb2xlcyI6W3siaWQiOiJURVJDRUlST1MiLCJkZXNjIjoiVGVyY2Vpcm9zIiwidHJhbnMiOlsiREVGQVVMVCJdfV0sImlzcyI6IlNBUEkiLCJjcGYiOiIwMTIyMzkzODcyNyIsIm5vbWUiOiJMRVZJIENBUlZBTEhPIENBTkVMTEEiLCJleHAiOjE3MzM3NjgyOTgsImlhdCI6MTczMzc2NTA1MywiZGVzYyI6IkNBTkVMTEEgRSBTQU5UT1MgQ09OVEFCSUxJREFERSBMVERBIn0.att4Wbm_ILbYfG4mUCXmE3msB7PmPpfBtzvhIZXGYdh2IXIEKCEC54TweXDgPhMN8hQtt7twIjUPkFmjN1r7iIWc99SMxU_Kwp_f5PrFql2uExtfKOpyhjVApjsWKcd2z524zxjm2mQZ3pUFgyrVgnhGhkiyGVh16bR6EUSJet-vWkGQYjI3badaXXvhbl_WLc52ZZNKdt03uBEhPhZdgdud2rV_gE9Gz3mZYXNs63uLkulyxrscmxmpb5TpN50QG61CbJVgH0kXHFpMtyjGgNVyNYMPY40pmC56FK5NV8T5SNqZlIjJA8D2lh9536LHa_L5GUpFci1ELTr2eR1uk2vqDsVmY173ort8wzYYxMfezpvUhd6CYDJDuzoOP3io32RTIfL8la560tEek95H8aaAFO83hsoaBc5hwOexg-ETCTIDUkWzO2DDsnTcysQ7k-cHjO7eptFiQrfLlc5CRdoMp6k9Kiu7pNnztghI_I1zJX-Wq4_UazOqDUXsheWljBsv8PXM3vo0nCdmAXB4MIl1jIVklpHgBpRWRbpzIeCVy_7AeAsDDhtfbN90SgFp1IYS_R9RI2aLBlRYPLbEemMIKGJSgOv7tBHUI9a27lxAGLnMCU285a1Bmb5g6ra2ktZ-XNLj5bB1pD0i7CVZENuZJZ_vIq6o4crWKKeXTH8'; // Substitua pelo token JWT real

// Lista de números de contribuintes
const contribuintes = [
    "14959642000127",
    "44464556000104",
    "42764132000130",
    "48262405000124",
    "48587498000167",
    "48778248000104",
    "48461235000107",
    "48721659000163",
    "48651078000100",
    "48618816000100",
    "41812438000151",
    "49744557000126",
    
];

function saveProtocolsToFile(protocols) {
  const protocolsFilePath = path.join(__dirname, 'protocolosLista.js');
  const errorFilePath = path.join(__dirname, 'protocolosErro.json');
  
  try {
    // Verifica se o arquivo já existe
    let existingProtocols = [];
    if (fs.existsSync(protocolsFilePath)) {
      const fileContent = fs.readFileSync(protocolsFilePath, 'utf8');
      const match = fileContent.match(/const protocolos = \[(.*?)\];/s);
      if (match && match[1].trim()) {
        existingProtocols = match[1]
          .split(',')
          .map(item => item.trim().replace(/['"]/g, ''));
      }
    }

    // Adiciona protocolos novos sem duplicatas
    const newProtocols = protocols
      .filter(p => p.protocolo && !existingProtocols.includes(p.protocolo))
      .map(p => `'${p.protocolo}'`);

    if (newProtocols.length > 0) {
      // Atualiza o arquivo com os protocolos novos
      const updatedProtocols = [...existingProtocols, ...newProtocols];
      const formattedProtocols = `const protocolos = [\n  ${updatedProtocols.join(',\n  ')}\n];`;
      fs.writeFileSync(protocolsFilePath, formattedProtocols, { flag: 'w' });
      console.log('Protocolos salvos com sucesso.');
    }

    // Salva a lista de contribuintes sem erros
    const successfulContributors = protocols
      .filter(p => p.protocolo && !p.error)
      .map(p => `'${p.contribuinte}'`);
    if (successfulContributors.length > 0) {
      const formattedContributors = `\nconst contribuintesSemErro = [\n  ${successfulContributors.join(',\n  ')}\n];`;
      fs.appendFileSync(protocolsFilePath, formattedContributors);
      console.log('Contribuintes sem erro salvos com sucesso.');
    }

    // Salva as entradas com erro em um arquivo separado
    const errorEntries = protocols.filter(p => p.error);
    if (errorEntries.length > 0) {
      fs.writeFileSync(errorFilePath, JSON.stringify(errorEntries, null, 2));
      console.log('Erros salvos com sucesso.');
    }
  } catch (error) {
    console.error('Erro ao salvar protocolos:', error.message);
  }
}


// Função para fazer a requisição
async function makeRequest(contribuinte) {
  const requestBody = {
    contratante: {
      numero: "06310711000149",
      tipo: 2,
    },
    autorPedidoDados: {
      numero: "06310711000149",
      tipo: 2,
    },
    contribuinte: {
      numero: contribuinte,
      tipo: 2,
    },
    pedidoDados: {
      idSistema: "SITFIS",
      idServico: "SOLICITARPROTOCOLO91",
      versaoSistema: "2.0",
      dados: "",
    },
  };

  try {
    // Faz a requisição POST
    const response = await axios.post(API_URL, requestBody, {
      headers: {
        Authorization: `Bearer ${ACCESS_TOKEN}`,
        "Content-Type": "application/json",
        jwt_token: JWT_TOKEN,
      },
    });

    // Extrai o número do protocolo da resposta
    const protocolo = response.data?.dados;
    console.log(`Protocolo obtido para contribuinte ${contribuinte}: ${protocolo}`);
    return { contribuinte, protocolo, error: null }; // Nenhum erro
  } catch (error) {
    console.error(`Erro ao processar contribuinte ${contribuinte}:`, error.response?.data || error.message);
    return { contribuinte, protocolo: null, error: error.message }; // Inclui a mensagem de erro
  }
}

// Função para processar todas as requisições
async function processRequests() {
  const protocols = [];

  for (const contribuinte of contribuintes) {
    const result = await makeRequest(contribuinte);
    protocols.push(result);

    // Atraso opcional entre as requisições para evitar sobrecarga
    await new Promise(resolve => setTimeout(resolve, 2000));
  }

  // Salva todos os protocolos no arquivo
  saveProtocolsToFile(protocols);
}

processRequests();
