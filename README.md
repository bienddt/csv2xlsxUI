# CSV para XLSX / ODS

Aplicação desktop em Python com interface Tkinter para converter CSV em planilhas Excel (`.xlsx`) ou OpenDocument (`.ods`, LibreOffice).

## Executar

Requer Python 3.10 ou superior com Tkinter. No Windows, instale o suporte Tcl/Tk no instalador do Python.

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python main.py
```

No Linux/macOS, ative com `source .venv/bin/activate`. Algumas distribuições Linux exigem o pacote `python3-tk`.

## Gerar executável Windows

Execute no PowerShell, com Python 3.10 ou superior instalado:

```powershell
.\build.ps1
```

O script prepara `.venv`, instala as dependências de build e gera `dist\CSV2XLSX.exe` com PyInstaller. O executável inclui Python e Tkinter, abre sem console e pode ser executado em Windows sem instalar Python. Gere o build no Windows da arquitetura de destino.

## Uso

1. Clique em **Selecionar CSV**. A aplicação analisa o arquivo e mostra até 20 registros.
2. Escolha **XLSX** ou **ODS**.
3. Confira os acentos e os campos na prévia. Se necessário, escolha manualmente a codificação, o delimitador ou o tipo de aspas e clique em **Analisar / atualizar prévia**.
4. Clique em **Converter e salvar** e selecione o destino. Arquivos existentes exigem confirmação de substituição.

## Detecção e preservação

- Detecta vírgula, ponto e vírgula, tabulação, barra vertical e dois-pontos como separadores.
- Reconhece aspas simples ou duplas como delimitador de texto, incluindo separadores e quebras de linha dentro de campos. Aspas escapadas seguem a convenção CSV de duplicação (`""` ou `''`).
- Cada CSV usa um único tipo de aspas para delimitar campos; aspas do outro tipo são preservadas como conteúdo. Arquivos que misturam as duas convenções de delimitação devem ser normalizados antes da conversão.
- Prioriza BOM e UTF-8 e usa análise estatística para outras codificações, com preferência por Windows-1252 em casos ambíguos. Detecção é uma estimativa: confira a prévia e ajuste manualmente quando necessário. Não repara texto que já foi salvo com acentos corrompidos.
- Grava texto Unicode no formato de destino. Preserva zeros iniciais, datas e números como texto para evitar alteração dos dados. Conteúdos iniciados por `=` permanecem texto, sem executar fórmulas.
- Não presume cabeçalho: todos os registros são exportados na aba **Dados**. Não aplica formatação numérica nem converte tipos automaticamente.
- Executa a leitura e a exportação em uma thread para manter a interface responsiva. O arquivo inteiro é carregado em memória, portanto o uso de RAM cresce com o CSV. A prévia também analisa o arquivo inteiro.
- Valida os limites de linhas, colunas e tamanho de célula do XLSX. A gravação usa um arquivo temporário no destino, substituído somente após a exportação terminar.

## Estrutura e testes

- `main.py`: interface gráfica.
- `converter.py`: leitura, detecção e exportação, utilizável sem interface.
- `tests/test_converter.py`: testes de detecção e integridade dos formatos gerados.
- `tests/test_e2e.py`: fluxos completos pela interface Tkinter real, acionando botões, opções, thread de processamento e verificando os arquivos produzidos.

```powershell
python -m unittest discover -s tests -v
```

Para executar somente os testes E2E:

```powershell
python -m unittest discover -s tests -p test_e2e.py -v
```

Os testes E2E exigem uma sessão gráfica com Tcl/Tk e abrem e fecham as janelas automaticamente. Em Linux sem display, use `xvfb-run -a python -m unittest discover -s tests -v` após instalar Tk e Xvfb. A ausência de display gera falha, não um teste ignorado.

Os diálogos nativos de abrir/salvar e as caixas de mensagem são substituídos por respostas controladas, sem exigir cliques humanos. O restante do fluxo é real: widgets Tkinter, fila de eventos, leitura CSV, detecção de formato, exportação e releitura independente dos arquivos. A suíte não valida a aparência visual nem os cliques nas janelas nativas do sistema operacional.

Cobertura dos fluxos E2E:

- CSV UTF-8 com BOM para XLSX, incluindo acentos, zeros iniciais, campos multilinha e fórmulas preservadas como texto.
- CSV Windows-1252 com aspas simples para ODS.
- Ajustes manuais de leitura e atualização da prévia.
- Exportação de todos os registros, mesmo quando a prévia mostra somente 20.
- Cancelamento de seleção e salvamento e tentativa de conversão sem arquivo.
- Extensão incompatível sem sobrescrever o destino.
- Recuperação após erro de leitura e nova conversão bem-sucedida.
- Interface responsiva e controles desabilitados durante o processamento.

O workflow `.github/workflows/tests.yml` executa a suíte no Windows com Python 3.10 e 3.13 em pushes, pull requests ou por acionamento manual no GitHub Actions.

## Licença

GNU General Public License v3.0 (GPL-3.0). Consulte [LICENSE](LICENSE).
