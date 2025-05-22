📊 Automated Excel Report
Automated Excel Report é um script em Python que automatiza a extração, transformação e geração de relatórios a partir de planilhas Excel. Ele lê dados específicos de um arquivo .xlsx, organiza as informações em um novo formato, aplica alinhamento centralizado em colunas selecionadas e salva o resultado final em uma nova planilha.

✅ Funcionalidades
📥 Extração automática de campos importantes:

Data de Atendimento
Cliente
Início
Término
Tempo de Execução
SLA
Motivo da Visita
Reparo Realizado

📝 Geração de nova planilha com dados organizados e cabeçalhos personalizados.

🎯 Alinhamento centralizado em colunas específicas (A, C, D, E, F) para melhor visualização.

💾 Salvamento automático do arquivo de saída.

🚀 Tecnologias Utilizadas
Python 3.x
OpenPyXL

🛠️ Como Usar
Clone o repositório:

bash
Copiar
Editar
git clone https://github.com/seuusuario/automated-excel-report.git
cd automated-excel-report
Instale as dependências:

bash
Copiar
Editar
pip install openpyxl

Configure o caminho do arquivo de origem
No script excel_report_generator.py, altere a variável:

python
Copiar
Editar
caminho_origem = r"C:\Users\SeuUsuario\Desktop\Pasta2.xlsx"
Execute o script:

bash
Copiar
Editar
python excel_report_generator.py
Resultado:
O arquivo resultado_final.xlsx será gerado automaticamente na pasta configurada.

🗂️ Estrutura do Projeto
bash
Copiar
Editar
automated-excel-report/

├── excel_report_generator.py

├── README.md

└── resultado_final.xlsx  # (gerado após a execução)

🤝 Contribuição
Contribuições são bem-vindas!
Sinta-se à vontade para abrir issues ou enviar pull requests.

📄 Licença
Este projeto está licenciado sob a licença MIT.
Consulte o arquivo LICENSE para mais detalhes.

⭐️ Dê uma estrela!
Se você gostou deste projeto, deixe uma ⭐️ para ajudar a divulgá-lo!

📬 Contato
LinkedIn: ronny-luiz-dos-santos

Email: ronny32luiz@gmail.com

🏷️ Tags
#python #excel #automation #openpyxl #report
