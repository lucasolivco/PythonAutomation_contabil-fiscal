const fs = require('fs');
const path = require('path');
const axios = require('axios');
const pdfParse = require('pdf-parse');

// Configurações
const API_URL = 'https://gateway.apiserpro.serpro.gov.br/integra-contador/v1/Emitir';
const ACCESS_TOKEN = 'eyJ4NXQiOiJNalJqWkRRMU1EQmtPR1JqWW1Jek9EVmxaRFEzWkdFeU1EVTVabU5rWldVeU9XUmhPRFZpTnciLCJraWQiOiJNRGN5WVdFNU16ZzVaVFJrT1dVME1XWTBNVE0xTkdJMllqbG1OVFZrT0dJd01UVmlORGRpWldJM1pEUmpZVEpsTTJaa05Ua3dNR1F3TWpZeVlXTmxNQV9SUzI1NiIsImFsZyI6IlJTMjU2In0.eyJzdWIiOiJhdXRlbnRpa3VzIiwiYXV0IjoiQVBQTElDQVRJT04iLCJhdWQiOiJPVEpqd1dscnJSNFJmRENDNXBKMTV4X09KaGNhIiwibmJmIjoxNzMzNzcwNTU3LCJhenAiOiJPVEpqd1dscnJSNFJmRENDNXBKMTV4X09KaGNhIiwic2NvcGUiOiJkZWZhdWx0IiwiaXNzIjoiaHR0cHM6XC9cL3B1Ymxpc2hlci5hcGlzZXJwcm8uc2VycHJvLmdvdi5icjo0NDNcL29hdXRoMlwvdG9rZW4iLCJyZWFsbSI6eyJzaWduaW5nX3RlbmFudCI6ImNhcmJvbi5zdXBlciJ9LCJleHAiOjE3MzM3NzQxNTcsImlhdCI6MTczMzc3MDU1NywianRpIjoiZWM1ZDMyOWYtNTJhOS00MGIwLWFhZDItZWRmYTJhNWZlYjFhIn0.d0vAXibPIZseDizZyDbaXIabwhxKc_OCPWzc9BbgWvVq7jqHa6suesg_CF-w2EMIiXfwEDeUqdGZS26VTSzH-PH96NV01reYAhBCpEBqyyUNPAkxxDtm4qtAN20Y64t_i3ZemOsJOEzSugVs_Fbn7lvJ0k52PCcAToVOLbRYK6PN0vVyAy2GQgzIiMkU-Ky3PFAJm0ILSIs24tK5nqf9OHhFoa8hbPPMoDRGl2-nx4_1uJcvU6Q31V3l8vodXGuGBxV8HYngImGFEhKz2WPvIPIGLBpH0oVPeW5OnCPq7HfMEapAOz8w9l4EjEK-Lkuo5KiAoytSwfjmrqm_rPMmAA'; // Substitua pelo token de acesso real
const JWT_TOKEN = 'eyJhbGciOiJSUzI1NiJ9.eyJzdWIiOiIwNjMxMDcxMTAwMDE0OSIsInRpcG8iOiJQSiIsImFjY2VzcyI6ImV5SjROWFFpT2lKTmFsSnFXa1JSTVUxRVFtdFBSMUpxV1cxSmVrOUVWbXhhUkZFeldrZEZlVTFFVlRWYWJVNXJXbGRWZVU5WFVtaFBSRlpwVG5jaUxDSnJhV1FpT2lKTlJHTjVXVmRGTlUxNlp6VmFWRkpyVDFkVk1FMVhXVEJOVkUweFRrZEpNbGxxYkcxT1ZGWnJUMGRKZDAxVVZtbE9SR1JwV2xkSk0xcEVVbXBaVkVwc1RUSmFhMDVVYTNkTlIxRjNUV3BaZVZsWFRteE5RVjlTVXpJMU5pSXNJbUZzWnlJNklsSlRNalUySW4wLmV5SnpkV0lpT2lKaGRYUmxiblJwYTNWeklpd2lZWFYwSWpvaVFWQlFURWxEUVZSSlQwNGlMQ0poZFdRaU9pSlBWRXBxZDFkc2NuSlNORkptUkVORE5YQktNVFY0WDA5S2FHTmhJaXdpYm1KbUlqb3hOek16Tnpjd05UVTNMQ0poZW5BaU9pSlBWRXBxZDFkc2NuSlNORkptUkVORE5YQktNVFY0WDA5S2FHTmhJaXdpYzJOdmNHVWlPaUprWldaaGRXeDBJaXdpYVhOeklqb2lhSFIwY0hNNlhDOWNMM0IxWW14cGMyaGxjaTVoY0dselpYSndjbTh1YzJWeWNISnZMbWR2ZGk1aWNqbzBORE5jTDI5aGRYUm9NbHd2ZEc5clpXNGlMQ0p5WldGc2JTSTZleUp6YVdkdWFXNW5YM1JsYm1GdWRDSTZJbU5oY21KdmJpNXpkWEJsY2lKOUxDSmxlSEFpT2pFM016TTNOelF4TlRjc0ltbGhkQ0k2TVRjek16YzNNRFUxTnl3aWFuUnBJam9pWldNMVpETXlPV1l0TlRKaE9TMDBNR0l3TFdGaFpESXRaV1JtWVRKaE5XWmxZakZoSW4wLmQwdkFYaWJQSVpzZURpelp5RGJhWElhYndoeEtjX09DUFd6YzlCYmdXdlZxN2pxSGE2c3Vlc2dfQ0YtdzJFTUlpWGZ3RURlVXFkR1pTMjZWVFN6SC1QSDk2TlYwMXJlWUFoQkNwRUJxeXlVTlBBa3h4RHRtNHF0QU4yMFk2NHRfaTNaZW1Pc0pPRXpTdWdWc19GYm43bHZKMGs1MlBDY0FUb1ZPTGJSWUs2UE4wdlZ5QXkyR1FneklpTWtVLUt5M1BGQUptMElMU0lzMjR0SzVucWY5T0hoRm9hOGhiUFBNb0RSR2wyLW54NF8xdUpjdlU2UTMxVjNsOHZvZFhHdUdCeFY4SFluZ0ltR0ZFaEt6MldQdklQSUdMQnBIMG9WUGVXNU9uQ1BxN0hmTUVhcEFPejh3OWw0RWpFSy1Ma3VvNUtpQW95dFN3ZmptcnFtX3JQTW1BQSIsImNlcnRpZmljYXRlRGV0YWlscyI6eyJzZXJpYWxOdW1iZXIiOiI4MDg2ODUyNTU0MjI1OTg2MjM3IiwiaXNzdWVyRE4iOiJDTj1BQyBTQUZFV0VCIFJGQiB2NSwgT1U9U2VjcmV0YXJpYSBkYSBSZWNlaXRhIEZlZGVyYWwgZG8gQnJhc2lsIC0gUkZCLCBPPUlDUC1CcmFzaWwsIEM9QlIiLCJ2YWxpZFVudGlsIjoiMjAyNS0xMC0wMVQxNTowMToxMS0wMzAwIiwidmFsaWRGcm9tIjoiMjAyNC0xMC0wMVQxNTowMToxMS0wMzAwIiwidmVyc2lvbiI6IjMiLCJzdWJqZWN0RE4iOiJDTj1DQU5FTExBIEUgU0FOVE9TIENPTlRBQklMSURBREUgTFREQTowNjMxMDcxMTAwMDE0OSwgT1U9dmlkZW9jb25mZXJlbmNpYSwgT1U9NDg5MTU2NTAwMDAxOTMsIE9VPVJGQiBlLUNOUEogQTEsIE9VPVNlY3JldGFyaWEgZGEgUmVjZWl0YSBGZWRlcmFsIGRvIEJyYXNpbCAtIFJGQiwgTD1WT0xUQSBSRURPTkRBLCBTVD1SSiwgTz1JQ1AtQnJhc2lsLCBDPUJSIn0sImlwIjoiMTc3LjEzMS4xODUuMTAiLCJyb2xlcyI6W3siaWQiOiJURVJDRUlST1MiLCJkZXNjIjoiVGVyY2Vpcm9zIiwidHJhbnMiOlsiREVGQVVMVCJdfV0sImlzcyI6IlNBUEkiLCJjcGYiOiIwMTIyMzkzODcyNyIsIm5vbWUiOiJMRVZJIENBUlZBTEhPIENBTkVMTEEiLCJleHAiOjE3MzM3NzQxNTcsImlhdCI6MTczMzc3MTc2MCwiZGVzYyI6IkNBTkVMTEEgRSBTQU5UT1MgQ09OVEFCSUxJREFERSBMVERBIn0.V4IgW7-8YAVw5saRpcOYx7c_NDznX87N5qkJWNPLmTtlrfhBkRfn_7hXDzBtmAuOrtjr8aH9-8k4X_nko2YHc-Jl9dzSOxRKTauBd4OwYUVw8RpQuQ_4ff2q3K6Ju5Xf63vRp2dAppFkF8HKvhdgMQ0cE0zFXwQVBM6DQoFNbO-xlu_JicOy0q39-VArf1ErU1uNM6ZgmJPcYm0tUz76QK0nPZohEIup_PeN7zrfnonTgm8AkwNUczQPMxs61StjBEblettZlHJgijrIaBKgX3lTO1RUNDngxx2rGnY-r0ZXtPiVfXW-GoY74OSC8Ly2adOSHpHDQ99Rr4sUsgYo3ICZYpCX5QLBuuuSj9o-UDOzYj2iYlBxQXUxkrNgYKBtzpTQEIBYBmHjhNVbfRVSPil86ARywPgye8E9ctGhG0Nx4gr08yjzPm-nmZ0hP6XJvSRxgjiEux6NQkBd0TzLBwZP2D_ruf1N_vykPbfeWKv4j2hSpmW5t5HgKXFmzNjp9A5PvMlFzwTYNxqe7owjrPWz6-fhEE0Yr0p5hA8LBDVN_JZ8eP3XqgDXJXffVG65gheNp9qyCS_SXL1rFcUIT0xEmOChF_k1djI6ED2dB93nqZrVDQEtyXZQ0TKVohexJDfNvvntmwKRfM9UYhAG44rBoxDXnKDNq-UqxSLYZgk'; // Substitua pelo token JWT real

// Pastas de saída
const pdfDir = path.join(__dirname, 'pdfs');
const jsonDir = path.join(__dirname, 'json_responses');
const logsDir = path.join(__dirname, 'logs');

/// Garante que as pastas existam
if (!fs.existsSync(pdfDir)) fs.mkdirSync(pdfDir);
if (!fs.existsSync(jsonDir)) fs.mkdirSync(jsonDir);
if (!fs.existsSync(logsDir)) fs.mkdirSync(logsDir);

// Arquivos de log
const successLogPath = path.join(logsDir, 'success.log');
const errorLogPath = path.join(logsDir, 'errors.log');

// Função para gravar logs
function writeLog(filePath, message) {
  const now = new Date();
  const timestamp = now.toLocaleString('pt-BR'); // Formato local brasileiro
  fs.appendFileSync(filePath, `[${timestamp}] ${message}\n`);
}


// Valores dinâmicos para as requisições
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

]; // Adicione mais números de contribuinte aqui

const protocolos = [
  '{"protocoloRelatorio":"+S7N6c04XNZUVzmxWT7SzpkZA4xeDQC9p2YabqGhXSCQ02x6gUEBspFyPPMJip33i/YJ5BJ4Pj5LhDx533HHJQaZu/FC+G1pYOtTMYqokKY9DNa7ejodsLWTXbTXPWsFQV5UOBwFRK2GSe28B5Ev25jDnzpvVJPhg/Msm1EoSyMowm9Da8FPwik4O9O4I4ba6F8avRRBJhQKXBhmBpKqf/qG6/ZJ7IVuKITCQqQMP6nDIr3OJjSAVA==","tempoEspera":4000}',
  '{"protocoloRelatorio":"+S7N6c04XNZUVzmxWT7SzpkZA4xeDQC9GS+CtfrgFgCtIs6PWc9BtOnRGgeTptsPi/YJ5BJ4Pj5LhDx533HHJQaZu/FC+G1pYOtTMYqokKY9DNa7ejodsLWTXbTXPWsFfrC+HI81qOWGSe28B5Ev25jDnzpvVJPhg/Msm1EoSyPSr8MraPCcbKpjfWmu4wAV6F8avRRBJhQKXBhmBpKqf/qG6/ZJ7IVuKITCQqQMP6nDIr3OJjSAVA==","tempoEspera":4000}',
  '{"protocoloRelatorio":"+S7N6c04XNZUVzmxWT7SzpkZA4xeDQC9Lr8W04M9I+8CdcGm3ljhTOnRGgeTptsPi/YJ5BJ4Pj5LhDx533HHJQaZu/FC+G1pYOtTMYqokKY9DNa7ejodsLWTXbTXPWsFJByvDypNHyeGSe28B5Ev25jDnzpvVJPhg/Msm1EoSyMQ/LxfJl+1ycDC4VoFkzOt6F8avRRBJhQKXBhmBpKqf/qG6/ZJ7IVuKITCQqQMP6nDIr3OJjSAVA==","tempoEspera":4000}'
]; // Adicione mais protocolos correspondentes aos contribuintes

async function extractNomeEmpresaFromPDF(pdfBuffer) {
  try {
    const pdfData = await pdfParse(pdfBuffer);
    const text = pdfData.text;

    // Expressão regular para capturar o nome da empresa
    const nomeEmpresaPattern = /CNPJ:\s*\d{2}\.\d{3}\.\d{3}\s*-\s*(.+)/i;
    const match = text.match(nomeEmpresaPattern);

    if (match && match[1]) {
      return match[1].trim(); // Retorna o nome da empresa
    }

    return null; // Retorna null se não encontrar o campo
  } catch (error) {
    console.error("Erro ao extrair o nome da empresa do PDF:", error);
    return null;
  }
}

function sanitizeFileName(fileName) {
  return fileName.replace(/[<>:"/\\|?*]+/g, '').trim();
}

// Função para lidar com nomes duplicados
function getUniqueFileName(directory, baseName, extension) {
  let uniqueName = `${baseName}${extension}`;
  let counter = 1;

  while (fs.existsSync(path.join(directory, uniqueName))) {
    uniqueName = `${baseName} (${counter})${extension}`;
    counter++;
  }

  return uniqueName;
}

async function processRequests() {
  for (let i = 0; i < contribuintes.length; i++) {
    const numeroContribuinte = contribuintes[i];
    const protocoloRelatorio = protocolos[i];
    const requestId = `request_${i + 1}`; // Identificador único para os arquivos

    // Monta o corpo da requisição
    const requestBody = {
      contratante: { numero: "06310711000149", tipo: 2 },
      autorPedidoDados: { numero: "06310711000149", tipo: 2 },
      contribuinte: { numero: numeroContribuinte, tipo: 2 },
      pedidoDados: {
        idSistema: "SITFIS",
        idServico: "RELATORIOSITFIS92",
        versaoSistema: "2.0",
        dados: protocoloRelatorio,
      },
    };

    try {
      // Faz a requisição para a API
      const response = await axios.post(API_URL, requestBody, {
        headers: {
          Authorization: `Bearer ${ACCESS_TOKEN}`,
          "Content-Type": "application/json",
          jwt_token: JWT_TOKEN,
        },
      });

      const responseData = response.data;

      /// Salva a resposta JSON
      const jsonFilePath = path.join(jsonDir, `${requestId}.json`);
      fs.writeFileSync(jsonFilePath, JSON.stringify(responseData, null, 2));
      writeLog(successLogPath, `Resposta salva para ${requestId}`);

      // Decodifica e salva o PDF
      if (responseData.dados) {
        const dados = JSON.parse(responseData.dados);
        if (dados.pdf) {
          const pdfBuffer = Buffer.from(dados.pdf, "base64");

          // Extrai o nome da empresa a partir do PDF
          const nomeEmpresa = await extractNomeEmpresaFromPDF(pdfBuffer);

          // Define o nome do arquivo, lidando com duplicatas
          const sanitizedNomeEmpresa = nomeEmpresa ? sanitizeFileName(nomeEmpresa) : requestId;
          const baseFileName = sanitizedNomeEmpresa || `request_${i + 1}`;
          const pdfFileName = getUniqueFileName(pdfDir, baseFileName, ".pdf");
          const pdfFilePath = path.join(pdfDir, pdfFileName);

          // Salva o arquivo PDF
          fs.writeFileSync(pdfFilePath, pdfBuffer);

          // Log simplificado
          writeLog(successLogPath, `PDF salvo: ${pdfFileName}`);
        } else {
          writeLog(errorLogPath, `Campo "pdf" não encontrado para ${requestId}`);
        }
      } else {
        writeLog(errorLogPath, `Campo "dados" não encontrado para ${requestId}`);
      }
    } catch (error) {
      const errorMessage = error.response?.data || error.message;
      writeLog(errorLogPath, `Erro ao processar ${requestId}: ${JSON.stringify(errorMessage, null, 2)}`);
    }
  }

  console.log(`Processamento concluído. Confira os logs em ${logsDir}`);
}

// Inicia o processamento
processRequests();