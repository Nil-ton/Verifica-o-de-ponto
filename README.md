# Sistema de Filtro de Planilha de Ponto de Serviço

Este é um sistema desenvolvido em Python especificamente para **facilitar o processo de registro de folhas de ponto** no meu trabalho. O objetivo principal do programa é otimizar o tempo gasto verificando se uma equipe já foi registrada na planilha de ponto, evitando duplicações e melhorando a eficiência nas minhas atividades diárias.

### Funcionalidades:
- **Carregar dados de uma planilha Excel**: O sistema carrega os dados do arquivo Excel que contém o controle de ponto dos colaboradores.
- **Filtragem por Data e MAT ENC**: Permite filtrar os registros pela data e pelo MAT ENC (Matrícula do encarregado), ajudando a encontrar rapidamente os registros desejados.
- **Verificação de equipes já registradas**: O sistema verifica se a equipe já foi registrada na planilha antes de adicionar novos dados, economizando tempo e evitando registros duplicados.
- **Exibição dos resultados filtrados**: Mostra os resultados em uma tabela e avisa caso algum campo obrigatório esteja vazio.
- **Interface gráfica simples e intuitiva**: Utiliza Tkinter para criar uma interface fácil de usar, adequada para o meu trabalho.

### Como Usar:
1. **Instalar dependências**:
   O sistema requer as bibliotecas `tkinter` e `pandas`. Caso ainda não as tenha instaladas, use o seguinte comando para instalá-las:

   ```bash
   pip install pandas openpyxl
   ```

2. **Preparar o arquivo Excel**:
   O arquivo Excel utilizado deve ter a seguinte estrutura de colunas:
   - `DATA`: Data do serviço (formato dd/mm/yyyy)
   - `MAT`: Matrícula do colaborador
   - `COLABORADOR`: Nome do colaborador
   - `Foi para Campo Gerar Produção?`: Informação se o colaborador foi para o campo
   - `Participou do DDS?`: Informação sobre participação no DDS
   - `MAT ENC`: Matrícula do encarregado
   - `ENCARREGADO`: Nome do encarregado
   - `PLACA`: Placa do veículo
   - `PREFIXO`: Prefixo do veículo
   - `BASE`: Base de operação

3. **Rodar o script**:
   O script pode ser executado diretamente no terminal ou no seu ambiente de desenvolvimento Python:

   ```bash
   python filtro_planilha.py
   ```

4. **Filtrar os dados**:
   Ao rodar o programa, a interface gráfica será aberta. Você pode:
   - **Inserir a data** no formato `dd/mm/yyyy`.
   - **Inserir o MAT ENC** (matrícula do encarregado).
   - **Clicar no botão "Filtrar Dados"** para aplicar o filtro.
   - **Visualizar os resultados** na tabela abaixo, onde será exibido o registro filtrado.

5. **Exibir resultados**:
   - Se houver registros correspondentes ao filtro, eles serão exibidos na tabela.
   - Se algum campo obrigatório estiver vazio, um alerta será mostrado informando quais campos precisam ser preenchidos.
   - Caso não haja resultados, uma mensagem de "Nenhum resultado encontrado" será exibida.

### Estrutura do Código:

- **FiltroPlanilha**: Classe responsável por carregar, filtrar e salvar os dados da planilha Excel.
- **AplicativoGUI**: Classe que implementa a interface gráfica utilizando Tkinter.
- **Funções de Formatação e Filtragem**: A função `formatar_data` garante que a data seja inserida corretamente, e a função `filtrar_dados` aplica o filtro na planilha com base na data e MAT ENC.
  
### Exemplo de Uso:

1. Abra a aplicação.
2. Insira a data (por exemplo, `20/01/2025`).
3. Insira a matrícula do encarregado (por exemplo, `12345`).
4. Clique em "Filtrar Dados" para ver os resultados.

### Personalização:
Este sistema foi desenvolvido **especificamente para o meu trabalho**. Ele resolve um problema específico que enfrento diariamente, permitindo que eu registre e controle mais rapidamente as folhas de ponto das equipes, além de garantir que não haja duplicações de registros.

### Melhorias Futuras:
- Adicionar funcionalidades para salvar registros diretamente na planilha após a filtragem.
- Melhorar a usabilidade com mais validações de entrada.
- Implementar o controle de versões de planilhas para garantir que não se perdam dados importantes.

### Contribuindo:
Este projeto foi desenvolvido para atender às minhas necessidades pessoais de trabalho. Portanto, não é um sistema genérico e pode não atender a outros casos de uso. No entanto, se alguém desejar contribuir ou fazer melhorias, fique à vontade para criar um fork e enviar pull requests.

---

Desenvolvido por: Nilton Oliveira  
Licença: MIT
