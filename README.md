Análise de Dados - Portal da Transparência

Sobre o Projeto

  Este projeto realiza a análise de dados das viagens públicas disponibilizadas pelo Portal da Transparência. 
  O script processa e consolida os dados, gerando um relatório em Excel e um gráfico interativo.

Funcionalidades

  Carregamento e tratamento dos dados brutos do Portal da Transparência
  Cálculo de despesas totais por cargo público
  Filtros para cargos mais relevantes
  Inserção de gráfico

Estrutura do Projeto

   ArchivesDA
  |-- 2025_Viagem.csv   # Arquivo bruto baixado do portal
  |-- Output
  │   |-- tabela_2025.xlsx   # Arquivo de saída com relatório e gráfico


Bibliotecas necessárias:

    pip install pandas matplotlib xlsxwriter

Execute o script Python:

    python script.py

  O relatório final estará disponível na pasta ArchivesDA/Output
