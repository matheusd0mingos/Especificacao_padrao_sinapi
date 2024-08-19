# Especificacao_padrao_sinapi
Este script do Google Apps Script automatiza a criação de documentos técnicos no padrão SINAPI usando dados de uma planilha do Google Sheets. O script gera um documento no Google Docs a partir de um modelo predefinido e preenche o conteúdo com base nas informações da planilha.

# Padrão de planilha para input
Exemplo:
![image](https://github.com/user-attachments/assets/3bf2a28a-8a5c-4ac3-9ddf-fa2c287ee94b)

#Planilha start
![image](https://github.com/user-attachments/assets/116e3744-c758-4034-98aa-9b7b8f824bbe)

# Funcionalidade

    Leitura de Dados: Obtém dados da planilha ativa no Google Sheets, incluindo o número inicial, o nome do arquivo e os dados das linhas.
    Criação do Documento: Cria uma cópia de um documento modelo existente no Google Drive e o renomeia conforme especificado.
    Preenchimento do Documento: Preenche o documento com informações extraídas da planilha, formatando e organizando o conteúdo conforme especificado.
    Salvar e Compartilhar: Salva e fecha o documento criado e atualiza a planilha com o link para o novo documento.

# Como Usar

    Configuração:
        Abra o Google Sheets onde você deseja usar o script.
        Crie uma aba chamada "Preencher Documento" e insira seus dados conforme as colunas esperadas.
        Crie uma aba chamada "Start" e insira o número inicial e o nome do arquivo nas células B5 e B6, respectivamente.

    Script:
        Abra o Editor de Scripts do Google Apps (Extensões > Apps Script).
        Copie e cole o código fornecido no Editor de Scripts.
        Atualize os IDs do template e da pasta de destino conforme necessário.

    Execução:
        Execute a função createNewGoogleDocs() a partir do Editor de Scripts para gerar o documento técnico.
        O link para o novo documento será atualizado na célula E9 da aba "Start".
