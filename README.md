# CPF-Consultation

Este código utiliza várias bibliotecas e algoritmos para automatizar o processo de consulta de dados em uma planilha do Google Sheets e interagir com um site usando Selenium. Aqui está um resumo das principais:

Bibliotecas:
requests: Para fazer solicitações HTTP.
selenium: Para automatizar interações com o navegador.
gspread: Para interagir com planilhas do Google Sheets.
oauth2client: Para autenticar com o Google e acessar os serviços.
pyautogui: Para realizar ações de automação no ambiente de desktop.

Algoritmos/Funcionalidades:
Envio de Mensagens Telegram: send_msg2() e delete_msg() enviam e excluem mensagens no Telegram respectivamente, utilizando a API do Telegram.
Autenticação no Google Sheets: Usa as credenciais do Google para autenticar e acessar as planilhas.
Consulta de Dados na Planilha: A função lphr() consulta dados na planilha e envia mensagens para o Telegram com informações relevantes.
Automação com Selenium: A função selenium() automatiza interações com um site, preenchendo campos de formulário, clicando em botões e obtendo resultados.
Atualização de Dados na Planilha: Atualiza células na planilha com os resultados obtidos.
Essencialmente, o código automatiza o processo de consulta de dados em uma planilha do Google Sheets e interação com um site externo para obter informações específicas.
