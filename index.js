const { Client } = require('whatsapp-web.js');
const qrcode = require('qrcode-terminal');
const fs = require('fs');
const xlsx = require('xlsx');

// Objeto para armazenar a quantidade de mensagens por grupo ou contato
const mensagensEnviadas = {};
let totalMensagensDia = 0;  // Contador de mensagens enviadas no dia
let primeiroEnvioDia = null;  // Data e hora do primeiro envio do dia
let ultimoEnvioDia = null;  // Data e hora do último envio do dia

// Inicializa o cliente do WhatsApp
const client = new Client();

// Gera e exibe o QR Code no console para autenticação
client.on('qr', (qr) => {
    qrcode.generate(qr, { small: true });
    console.log('Escaneie o QR Code acima para autenticar no WhatsApp.');
});

// Informa que o cliente foi autenticado com sucesso e está pronto para uso
client.on('ready', () => {
    console.log('O bot está pronto para uso!');
});

// Monitora todas as mensagens enviadas por você
client.on('message_create', async (message) => {
    // Verifica se a mensagem foi enviada por você
    if (message.fromMe) {
        try {
            // Obtenha o chat (pode ser grupo ou contato individual)
            const chat = await message.getChat();
            const dataEnvio = new Date(message.timestamp * 1000);
            const conteudoMensagem = message.body;
            const horario = dataEnvio.toLocaleTimeString();
            const data = dataEnvio.toLocaleDateString();

            let nomeChat = '';

            // Atualiza o contador total de mensagens no dia
            totalMensagensDia += 1;

            // Atualiza a data do primeiro e último envio no dia
            if (!primeiroEnvioDia || dataEnvio < primeiroEnvioDia) {
                primeiroEnvioDia = dataEnvio;
            }
            if (!ultimoEnvioDia || dataEnvio > ultimoEnvioDia) {
                ultimoEnvioDia = dataEnvio;
            }

            // Verifica se o chat é um grupo
            if (chat.isGroup) {
                nomeChat = chat.name;

                console.log(`Você enviou uma mensagem para o grupo "${nomeChat}": "${conteudoMensagem}" - Data: ${data} - Horário: ${horario}`);
                
                // Armazena a mensagem enviada para o grupo
                if (!mensagensEnviadas[nomeChat]) {
                    mensagensEnviadas[nomeChat] = {
                        totalMensagens: 0,
                        dataInicio: dataEnvio,
                        dataFim: dataEnvio,
                    };
                }

                mensagensEnviadas[nomeChat].totalMensagens += 1;
                mensagensEnviadas[nomeChat].dataFim = dataEnvio; // Atualiza a data da última mensagem
            } else {
                // Se não for um grupo, é um contato individual
                const contato = await message.getContact();
                nomeChat = contato.pushname || contato.number;

                console.log(`Você enviou uma mensagem para "${nomeChat}": "${conteudoMensagem}" - Data: ${data} - Horário: ${horario}`);

                // Armazena a mensagem enviada para o contato individual
                if (!mensagensEnviadas[nomeChat]) {
                    mensagensEnviadas[nomeChat] = {
                        totalMensagens: 0,
                        dataInicio: dataEnvio,
                        dataFim: dataEnvio,
                    };
                }

                mensagensEnviadas[nomeChat].totalMensagens += 1;
                mensagensEnviadas[nomeChat].dataFim = dataEnvio; // Atualiza a data da última mensagem
            }

        } catch (error) {
            console.error('Erro ao processar a mensagem:', error);
        }
    }
});

// Inicia o cliente do WhatsApp
client.initialize();

// Função para calcular o período total em horas e minutos
function calcularPeriodoTotal(primeiroEnvio, ultimoEnvio) {
    const diffMs = ultimoEnvio - primeiroEnvio;
    const diffHoras = Math.floor(diffMs / (1000 * 60 * 60));
    const diffMinutos = Math.floor((diffMs % (1000 * 60 * 60)) / (1000 * 60));
    return `${diffHoras} horas e ${diffMinutos} minutos`;
}

// Função para salvar os dados no Excel
function salvarMensagensNoExcel() {
    const dados = [['Nome do Grupo/Contato', 'Número de Mensagens Enviadas', 'Data', 'Período do Envio']];
    
    // Preenche os dados com base no objeto "mensagensEnviadas"
    for (const [nome, info] of Object.entries(mensagensEnviadas)) {
        const dataInicio = info.dataInicio.toLocaleDateString();
        const horarioInicio = info.dataInicio.toLocaleTimeString();
        const dataFim = info.dataFim.toLocaleDateString();
        const horarioFim = info.dataFim.toLocaleTimeString();
        const periodo = `${dataInicio} ${horarioInicio} - ${dataFim} ${horarioFim}`;
        dados.push([nome, info.totalMensagens, dataInicio, periodo]);
    }

    // Adiciona o total de mensagens enviadas no dia e o período total
    if (primeiroEnvioDia && ultimoEnvioDia) {
        const periodoTotal = calcularPeriodoTotal(primeiroEnvioDia, ultimoEnvioDia);
        dados.push(['TOTAL DO DIA', totalMensagensDia, primeiroEnvioDia.toLocaleDateString(), periodoTotal]);
    }

    // Cria uma nova planilha no Excel com os dados
    const wb = xlsx.utils.book_new();
    const ws = xlsx.utils.aoa_to_sheet(dados);
    xlsx.utils.book_append_sheet(wb, ws, 'Mensagens Enviadas');
    
    // Salva o arquivo Excel
    const nomeArquivo = 'mensagens_enviadas.xlsx';
    xlsx.writeFile(wb, nomeArquivo);

    console.log(`As mensagens foram salvas no arquivo Excel: ${nomeArquivo}`);
}

// Escuta o evento de encerramento do script (CTRL + C)
process.on('SIGINT', () => {
    console.log('\nEncerrando o script e salvando os dados...');
    salvarMensagensNoExcel();
    process.exit();
});
