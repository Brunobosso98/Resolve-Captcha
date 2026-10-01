# AI Context

## Directory Structure Map

- `sigiss_modais.py`: rotina compartilhada para fechar os avisos "Estou Ciente" antes do login, inclusive quando vários modais repetem o ID `btnCiente`.
- `exportar_notas_tomados.py`: exportação de notas de serviços tomados.
- `serv_prestado_simples_nacional.py`: exportação de notas e encerramento mensal de prestados do Simples Nacional.
- `servicos_prestados_demais.py`: exportação de notas e geração de PDF/encerramento de prestados das demais empresas.
- `servicos_tomados.py`: processamento mensal de serviços tomados.
- `servicos_prestados.py`: automação base do SIGISS para login com captcha, alteração de competência e rotinas de serviços prestados.
- `livro_mensal_prestados.py`: exportação do livro mensal de serviços prestados.
- `livro_mensal_tomados2.py`: exportação do livro mensal de serviços tomados.
- `livro_anual_prestados_xls.py`: nova automação para livro anual em Excel, reutilizando o fluxo de login/captcha do SIGISS e renomeando o download para o CNPJ.
- `Senha Municipio Itapira.xlsx`: planilha com empresa, usuário/CNPJ, senha e status de processamento.
- `livroFiscalAnual/`: pasta de saída dos downloads anuais em Excel.
- `Livro/`: pasta de saída dos livros mensais já exportados.
- `NOTAS/`: rotinas relacionadas à emissão e exportação de notas fiscais.

## Business Logic & Feature Map

- Login SIGISS:
  usa CNPJ em `Usuário`, senha da planilha e captcha OCR com Tesseract.
- Avisos antes do login:
  `exportar_notas_tomados.py`, `serv_prestado_simples_nacional.py`, `servicos_prestados_demais.py` e `servicos_tomados.py` chamam `sigiss_modais.fechar_avisos_login()`. A rotina busca todos os botões `btnCiente`, aguarda o fim das transições do Bootstrap (`bs.modal._isTransitioning`), clica nos visíveis e aguarda cada um desaparecer; após um fechamento, espera até 3 segundos pelo próximo aviso, com limite total de 20 segundos. O site abre `myModal3`, `myModal2` e `myModal` nessa ordem. Falhas no fechamento geram mensagens específicas e uma nova tentativa de acesso para a empresa, respeitando o limite de tentativas existente.
- Competência:
  os scripts alteram mês e ano no painel principal após autenticação.
- Livro mensal:
  scripts existentes exportam documentos mensais de prestados e tomados.
- Livro anual em Excel:
  `livro_anual_prestados_xls.py` fixa a competência em dezembro de 2025, executa `livroAnualP_xls()` no console da página e renomeia o arquivo baixado para o CNPJ da empresa.
- Status operacional:
  a coluna `Status Processo` da planilha registra falhas de login, ausência de download e sucesso do processamento.
