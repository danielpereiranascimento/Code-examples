# PHC Scripts

Repositório com scripts para automação de tarefas no PHC, incluindo:

- Verificação de stock
- Envio de documentos por email
- Importação de dados de Excel
- Funções auxiliares reutilizáveis

---

## Estrutura

### `stock_alert/`
- `stock_verification.prg`  
  Verifica se há stock suficiente para encomendas de cliente. Envia email de alerta se necessário.

### `document_emails/`
- `send_documents_email.prg`  
  Seleciona documentos do cliente, gera PDFs e envia email HTML com cabeçalho e rodapé gráficos.

### `excel_import/`
- `import_excel_fn.prg`  
  Importa dados de Excel para a tabela FN, mapeando taxas de IVA e matrículas.

- `import_excel_general.prg`  
  Importação genérica de Excel para cursor temporário, útil para outros tipos de dados.

### `utils/`
- `helpers.prg`  
  Funções auxiliares, como envio de email, barra de progresso, formatação de dados, etc.

---

## Como usar

1. Colocar os scripts na pasta correspondente.
2. Configurar os caminhos de email e pastas de saída no PHC.
3. Executar os scripts a partir do PHC conforme necessidade.
4. Para importações Excel, certificar-se que os ficheiros têm as colunas corretas.

---

## Notas

- Scripts testados no ambiente PHC com Excel e email via SMTP.
- Sempre criar backup antes de atualizar a base de dados com importações.
