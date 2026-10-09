## Sistema de análise de logs e criação de relatório em Python

O código desenvolve um sistema de análise de logs utilizando **Python**. Ele lê os registros presentes em um arquivo de texto, separa as informações por data, hora, tipo e mensagem, contabiliza as ocorrências e gera um relatório formatado no formato **DOCX**.

Para realizar essas tarefas, o programa utiliza as bibliotecas `re`, `datetime`, `pathlib` e `python-docx`.

### Leitura e manipulação de arquivos

A classe `Path`, da biblioteca `pathlib`, é utilizada para representar os caminhos dos arquivos. O atributo `Path(__file__).parent` identifica a pasta em que o código está armazenado.

O método `exists()` verifica se o arquivo de logs existe antes de iniciar o processamento. Caso o arquivo não seja encontrado, o programa exibe uma mensagem e encerra a função com o comando `return`.

A função `open()` abre o arquivo de logs com a codificação `UTF-8`. O método `read()` lê todo o conteúdo do arquivo para que os registros possam ser analisados.

### Extração dos registros

A função `re.findall()` utiliza uma expressão regular para localizar e separar as informações de cada registro. Os dados são divididos em:

- Data;
- Hora;
- Tipo do registro;
- Mensagem.

O método `append()` adiciona cada registro extraído a uma lista no formato de dicionário.

O método `upper()` transforma o tipo do registro em letras maiúsculas, enquanto `strip()` remove espaços desnecessários no início e no final das mensagens.

### Formatação de datas

O método `datetime.strptime()` converte uma data armazenada como texto para um objeto de data.

Em seguida, o método `strftime()` formata a data no padrão brasileiro, utilizando dia, mês e ano. Esse método também é utilizado para inserir a data e o horário de geração no relatório.

O método `datetime.now()` obtém a data e o horário atuais.

### Contagem e organização dos registros

O método `get()` consulta a quantidade de ocorrências de cada tipo de log no dicionário. Quando um tipo ainda não foi registrado, o valor inicial utilizado é zero.

O método `setdefault()` cria uma entrada no dicionário para cada data, caso ela ainda não exista. Dessa forma, o programa consegue organizar e contar os registros de cada dia.

As funções `len()`, `sum()` e `sorted()` também são utilizadas:

- `len()` calcula o total de registros;
- `sum()` soma as ocorrências de cada dia;
- `sorted()` organiza as datas e os tipos de log em ordem.

### Criação do documento Word

A classe `Document()` cria um novo documento do Word.

O método `add_heading()` adiciona títulos e seções ao relatório, enquanto `add_paragraph()` cria novos parágrafos.

O método `add_run()` adiciona trechos de texto dentro dos parágrafos e permite aplicar diferentes estilos, como fonte, tamanho, cor, negrito e itálico.

O método `add_table()` cria as tabelas do relatório. Já o método `add_row()` adiciona novas linhas às tabelas.

O método `add_page_break()` insere uma quebra de página antes da seção que apresenta os registros organizados por dia.

O método `save()` salva o documento criado no formato `.docx`.

### Formatação do relatório

O código utiliza os recursos `Pt()`, `Cm()` e `RGBColor.from_string()` para configurar o tamanho das fontes, a largura das colunas e as cores do documento.

O método `set()` configura propriedades específicas dos elementos XML utilizados na formatação do Word.

O método `append()` também é empregado na estrutura XML para adicionar bordas, cores de fundo e o número automático das páginas.

As constantes `WD_ALIGN_PARAGRAPH` e `WD_TABLE_ALIGNMENT` definem o alinhamento dos parágrafos, títulos, células e tabelas.

### Estruturas de controle

O código utiliza o laço `for` para percorrer os registros, contar as ocorrências e preencher as tabelas.

As estruturas `if` e `else` verificam situações como:

- Existência do arquivo de logs;
- Presença de registros extraídos;
- Existência de registros do tipo `ERROR`;
- Aplicação de cores alternadas nas tabelas;
- Definição de cores especiais para cada tipo de log.

O comando `return` encerra uma função quando o arquivo não é encontrado ou quando nenhum registro pode ser extraído.

A condição `if __name__ == "__main__"` garante que a função `criar_relatorio()` seja executada somente quando o arquivo Python for iniciado diretamente.

### Principais métodos e funções utilizados

- `Path()`: representa caminhos de arquivos e pastas;
- `exists()`: verifica se o arquivo existe;
- `open()`: abre o arquivo de logs;
- `read()`: lê o conteúdo do arquivo;
- `re.findall()`: extrai os dados com uma expressão regular;
- `append()`: adiciona elementos a listas e estruturas XML;
- `upper()`: converte textos para letras maiúsculas;
- `strip()`: remove espaços desnecessários;
- `datetime.strptime()`: converte texto em data;
- `datetime.now()`: obtém a data e o horário atuais;
- `strftime()`: formata datas e horários;
- `get()`: consulta valores em dicionários;
- `setdefault()`: cria valores padrão em dicionários;
- `len()`: retorna a quantidade de registros;
- `sum()`: soma os valores;
- `sorted()`: ordena os registros;
- `Document()`: cria o documento Word;
- `add_heading()`: adiciona títulos;
- `add_paragraph()`: adiciona parágrafos;
- `add_run()`: adiciona textos aos parágrafos;
- `add_table()`: cria tabelas;
- `add_row()`: adiciona linhas às tabelas;
- `add_page_break()`: insere uma quebra de página;
- `save()`: salva o relatório em formato DOCX.

## Conclusão

O programa combina leitura de arquivos, expressões regulares, manipulação de datas, listas, dicionários e criação de documentos Word. Com esses recursos, ele transforma os registros de um arquivo de log em um relatório organizado, contendo o total de ocorrências por tipo, os registros de erro e a quantidade de eventos registrados em cada dia.
