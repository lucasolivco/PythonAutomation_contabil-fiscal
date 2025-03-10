const axios = require('axios');
const https = require('https');
const fs = require('fs');
const path = require('path');

// Configurações
const CERT_PATH = path.join(__dirname, '102 - CANELLA E SANTOS - SENHA Canella1234 - V 01.10.2025.pfx'); // Caminho para o arquivo do certificado e-CNPJ
const CERT_PASSWORD = 'Canella1234'; // Senha do certificado
const AUTH_URL = 'https://autenticacao.sapi.serpro.gov.br/authenticate';
const CONSUMER_KEY = 'OTJjwWlrrR4RfDCC5pJ15x_OJhca';
const CONSUMER_SECRET = 'hrdI6XFRusefnccjNhMgqUw4kf8a';

async function getAuthToken() {
    try {
        // Cria o agente HTTPS com o certificado digital
        const httpsAgent = new https.Agent({
            pfx: fs.readFileSync(CERT_PATH),
            passphrase: CERT_PASSWORD,
        });

        // Codifica as credenciais em Base64
        const authHeader = Buffer.from(`${CONSUMER_KEY}:${CONSUMER_SECRET}`).toString('base64');

        // Faz a requisição POST para autenticação
        const response = await axios.post(
            AUTH_URL,
            new URLSearchParams({ grant_type: 'client_credentials' }), // Corpo da requisição
            {
                headers: {
                    'Authorization': `Basic ${authHeader}`,
                    'role-type': 'TERCEIROS',
                    'content-type': 'application/x-www-form-urlencoded',
                },
                httpsAgent, // Utiliza o agente com o certificado
            }
        );

        console.log('Token de Acesso Bearer:', response.data.access_token);
        console.log('Token JWT:', response.data.jwt_token);
        return response.data;
    } catch (error) {
        console.error('Erro ao obter token:', error.response?.data || error.message);
    }
}

// Executa a função
getAuthToken();

