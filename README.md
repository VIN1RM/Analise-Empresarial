📊 Sistema de Análise e Documentação de Desempenho

Desenvolvimento feito em Python que utiliza Tkinter, Pandas, Matplotlib, Seaborn e python-docx para processar dados de desempenho, gerar gráficos e criar relatórios detalhados em formato Word. 
Este sistema é ideal para avaliar o desempenho de alunos ou colaboradores em disciplinas, identificando áreas que necessitam de reforço e produzindo uma documentação visualmente rica.


🎯 Funcionalidades

- Entrada de Dados: Cole dados diretamente de planilhas Excel em uma interface Tkinter.
  
- Processamento Automático: Converte dados em um DataFrame Pandas, tratando valores numéricos e removendo caracteres indesejados (ex.: "%").
  
- Análise de Desempenho: Identifica alunos/collaboradores com notas abaixo de 80% em disciplinas específicas.
  
- Geração de Gráficos:
    Gráfico de pizza: Proporção de disciplinas com necessidade de reforço.
    Gráfico de barras: Média de notas por disciplina.
    Gráficos individuais: Desempenho por aluno/colaborador.
    Histogramas: Distribuição de notas por disciplina.
  
- Relatório em Word: Cria um documento com:
    Lista de alunos/collaboradores com disciplinas para reforço.
    Gráficos gerais e individuais.
    Relatório por disciplina com ranking de colaboradores.
  
- Interface Intuitiva: Janelas Tkinter para entrada de nome da empresa e dados, com validações e mensagens de erro.
  
- Organização de Arquivos: Salva relatórios e gráficos em uma pasta dedicada na Área de Trabalho.


📋 Pré-requisitos

  Python 3.x (recomenda-se 3.8 ou superior)
  Bibliotecas necessárias:
    tkinter (incluso no Python)
    pandas
    python-docx
    matplotlib
    seaborn


🎮 Como Usar

- Iniciar a Aplicação:
    Ao executar, uma janela pergunta se deseja iniciar uma nova análise.
    Clique em "Sim" para prosseguir.

- Inserir Nome da Empresa:
    Digite o nome da empresa na janela inicial e clique em "Confirmar".
    O nome será usado para nomear a pasta de saída.

- Colar Dados:
    Copie os dados de uma planilha Excel (com cabeçalhos e valores).
    Cole na área de texto da segunda janela e clique em "Confirmar".
  
- Saída Gerada:
    Uma pasta chamada Documentação <Nome da Empresa> será criada na Área de Trabalho.
    Dentro dela, você encontrará:
        relatorio.docx: Documento Word com análise e gráficos.
        Arquivos PNG: Gráficos gerais, individuais e por disciplina.
  
- Formato dos Dados:
    Os dados devem ter:
        Primeira coluna: Nomes dos alunos/collaboradores.
        Demais colunas: Disciplinas com notas (valores numéricos ou com "%").
  
    Exemplo:
  
        Nome    Matemática    Física    Química
        João    85%           70%       90%
        Maria   60%           80%       75%


🌟 Possíveis Melhorias

- Adicionar suporte para múltiplas empresas em uma única execução.



👨‍💻 Autor
Vinícius Rodrigues Martins
GitHub: VIN1RM
Entre em contato para dúvidas ou sugestões!
