Este código é um validador automatizado de números de WhatsApp que utiliza a biblioteca Selenium para automatizar interações com o WhatsApp Web e a biblioteca OpenPyXL para registrar os resultados em uma planilha Excel. O objetivo principal do script é validar se números de telefone possuem conta no WhatsApp.

Funcionamento:
O script abre o WhatsApp Web e aguarda que o usuário faça login manualmente.
Uma vez logado, ele clica em uma conversa específica que contém os números a serem validados. Para que funcione, é necessário inserir manualmente os números na conversa e obter o XPath específico tanto da conversa quanto de cada número a ser validado.
Para cada número encontrado, o script verifica se a opção "Conversar com (número)" aparece na interface do WhatsApp Web.
Se a opção for encontrada, o script registra "Whatsapp ok" na planilha Excel.
Caso contrário, registra "Sem whatsapp".
Todos os resultados são salvos em uma planilha Excel.
