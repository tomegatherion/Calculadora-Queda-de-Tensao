## Calculadora de Queda de Tensão

Esta é uma aplicação desenvolvida em Python utilizando a biblioteca PySide6 para criar uma interface gráfica de usuário (GUI). A calculadora permite calcular a queda de tensão em um circuito elétrico com base em diversos parâmetros de entrada fornecidos pelo usuário.

### Funcionalidades

- Cálculo da queda de tensão em circuitos monofásicos e trifásicos
- Exibição da potência aparente, corrente, tensão a jusante, queda de tensão parcial e acumulada
- Indicador visual da queda de tensão em relação ao limite especificado
- Histórico dos cálculos realizados
- Opção de salvar os resultados na tabela de histórico
- Exportação dos dados do histórico para um arquivo Excel
- Suporte a diferentes estilos visuais da interface

### Parâmetros de Entrada

- Nome do circuito
- Esquema de ligação (1FNT, 2FNT ou 3FNT)
- Tensão de entrada
- Tensão a montante
- Fator de potência
- Potência ativa
- Comprimento do condutor
- Número de trifólios
- Seção do condutor
- Material do condutor (cobre ou alumínio)
- Limite de queda de tensão (4% ou 7%)

### Resultados Calculados

- Potência aparente
- Corrente
- Tensão a jusante
- Queda de tensão parcial
- Queda de tensão acumulada
- Indicador visual da queda de tensão em relação ao limite especificado

### Como Usar

1. Execute o script `CalculoQuedaTensão.py` para iniciar a aplicação.
2. Preencha os parâmetros de entrada de acordo com o circuito desejado.
3. Clique no botão "Calcular" para obter os resultados.
4. Para salvar os resultados no histórico, clique no botão "Salvar".
5. Para exportar os dados do histórico para um arquivo Excel, clique no botão "Exportar".
6. Use os botões "Tema" e "Histórico" para alternar entre os estilos visuais e exibir/ocultar a tabela de histórico, respectivamente.

### Instalação das Dependências

Para instalar as dependências necessárias, siga os passos abaixo:

1. **Instale o Python 3.x**: Certifique-se de que você tem o Python 3.x instalado em seu sistema. Você pode baixá-lo em [python.org](https://www.python.org/downloads/).

2. **Crie um ambiente virtual (opcional)**: É recomendado criar um ambiente virtual para evitar conflitos de dependências. Você pode fazer isso com os seguintes comandos:
 
   ```bash
   python -m venv meu_ambiente
   source meu_ambiente/bin/activate  # No Linux ou Mac
   meu_ambiente\Scripts\activate  # No Windows
   ```
3. Instale as dependências: Execute o seguinte comando para instalar as bibliotecas necessárias:
    ```bash
    pip install -r requirements.txt
    ```
    
### Contribuição


Contribuições são bem-vindas! Se você encontrar algum problema ou tiver sugestões de melhorias, sinta-se à vontade para abrir uma issue ou enviar um pull request.
