# Dimensionador de Pallets

Aplicativo Streamlit para dimensionar pallets a partir do volume diário previsto, sem exigir conhecimento prévio dos destinos.

## Identidade visual

- logotipo e favicon institucionais da BBM Logística incorporados ao projeto;
- fonte Urbanist aplicada globalmente;
- paleta em azul-marinho e laranja;
- cabeçalhos, indicadores, abas, botões e navegação responsivos.
- datas exibidas em `dd/mm/aaaa` e números no padrão brasileiro;
- médias de volumes por pallet apresentadas como números inteiros;
- gráficos com destaque vermelho para o dimensionamento atual e rótulos
  seletivos nos pontos históricos relevantes.
- gráficos Plotly com zoom, movimentação, restauração da visualização, tooltip
  e download como imagem.

O logotipo foi obtido no site institucional:
`https://bbmlogistica.com.br/wp-content/uploads/2025/11/bbm-logistica-logo.png`.

## Funcionalidades

- painel diário central para consultar previsões salvas por data, mesmo antes
  de carregar a execução;
- entrada manual da data e dos volumes previstos;
- simulação histórica de vários dias por tabela editável ou upload de Excel/CSV;
- cálculo de pallets esperados, capacidade recomendada e cenário de alta densidade;
- seleção automática de dias comparáveis por dia da semana e faixa de volume;
- descrição detalhada e ajuda contextual para todos os parâmetros ajustáveis
  do modelo;
- página `Informações` com parâmetros padrão, lógica de seleção dos dias e
  interpretação dos resultados;
- maior peso para operações recentes;
- distribuição estimada por praça;
- média estimada de volumes por pallet específica para cada praça;
- segmentador de praça nos gráficos da simulação diária, com histórico e dias
  comparáveis específicos da unidade selecionada;
- filtro multisseleção por UF nas principais visões por praça, usando cadastro
  de unidades validado pelo CEP;
- upload incremental da base executada;
- deduplicação e atualização de carregamentos corrigidos;
- limpeza de espaços comuns, espaços especiais e caracteres invisíveis nos
  códigos das praças;
- consolidação automática de variações já existentes no banco, sem alterar o
  total de pallets;
- comparação entre previsão e execução;
- exclusão confirmada de todos os cenários previstos salvos;
- histórico persistido em SQLite.

Ao abrir o aplicativo, o `Painel diário` consulta inicialmente a data atual. Na
`Simulação diária`, a data também começa em hoje e os volumes previstos começam
em zero. O botão `Salvar previsão e abrir painel` grava o cenário e abre
diretamente o painel da data salva.

## Execução

1. Instale o Python 3.11 ou superior.
2. Abra o terminal na pasta do projeto.
3. Instale as dependências:

   ```bash
   pip install -r requirements.txt
   ```

4. Execute:

   ```bash
   streamlit run app.py
   ```

O aplicativo cria automaticamente `data/pallet_planner.db` e carrega `data/base_inicial.xlsx` na primeira execução.

O banco SQLite permanece na máquina onde o aplicativo estiver rodando. Em uma
hospedagem que recria o ambiente a cada publicação, conecte o aplicativo a um
banco persistente, como PostgreSQL, antes do uso contínuo em produção.

## Estrutura esperada para uploads

A planilha deve possuir:

- `Data de expedição`;
- `Volumes`;
- destinos nas colunas `1` a `5`;
- pallets nas colunas `P1` a `P5`.

CWB e CW2 são consolidados como `CWB/CW2`. FLI e FL2 são consolidados como `FLI/FL2`. Os demais agrupamentos permanecem separados.

## Regra de cálculo

1. Usa somente registros anteriores à data simulada.
2. Procura o mesmo dia da semana e volumes dentro de ±15%.
3. Se a amostra for pequena, amplia para ±25%.
4. Se ainda for insuficiente, utiliza todos os registros do mesmo dia da semana.
5. Calcula a densidade esperada ponderando mais as semanas recentes.
6. Recalcula como o modelo teria performado nos 14 dias anteriores, sempre sem
   usar informações futuras.
7. Aplica à estimativa um fator adaptativo baseado na relação entre pallets
   executados e pallets estimados, dando mais peso à última semana.
8. Calcula a capacidade recomendada acrescentando o percentil 80 do erro
   absoluto recente do modelo.
9. Estima a densidade de cada praça pelas cargas históricas em que ela apareceu,
   ponderando pallets e recência. Praças com pouca amostra são aproximadas da
   média geral.
10. Reconcilia os volumes das praças para que a soma seja exatamente igual ao
    volume total informado na simulação.

Os parâmetros podem ser ajustados dentro do aplicativo.

## Rotina diária sugerida

1. Informe a data e o total de volumes previstos.
2. Use `Pallets esperados` como referência e `Capacidade recomendada` como
   margem operacional.
3. Salve a previsão; o painel diário será aberto automaticamente.
4. Depois da expedição, carregue a base executada atualizada.
5. Consulte `Previsto × realizado` para acompanhar os desvios totais e por
   praça.

Para avaliar datas anteriores em lote, use `Simulação histórica`. A página
aceita as colunas `Data` e `Volumes planejados`, compara o cálculo com a
execução e permite baixar os resultados diários e por praça.
