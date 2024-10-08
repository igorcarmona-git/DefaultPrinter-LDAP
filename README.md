# Solução interna para um problema de seleção de impressoras padrão no Windows em diversos locais diferentes.

Trabalho em uma empresa onde os funcionários que atuam em determinada atividade não ficam somente no mesmo setor. Acontecia que, a pessoa selecionava a impressora como padrão em um local e quando se deslocava para um local diferente com outra impressora tinha que ficar selecionando a impressora toda vez como padrão. (O sistema só imprimia as guias de atendimento e etc, somente se ativo a opção, com isso, ocasionava diversos problemas).

**- Observações:**
- Entende-se de que você já tenha o servidor AD configurado, politicas de grupo definidas e um scritlogon definido para que todos os usuários na rede possam executar ao fazer login.
- Python 3.12 deve estar instalado no computador.

## Funcionalidades

Este código ele faz login no servidor LDAP (Lightweight Directory Access Protocol), verifica se o usuário logado tem impressoras disponíveis localmente naquela máquina e verifica se o hostname do computador está vinculado ao nome da impressora disponivel para o usuário dele. Se tiver impressora, ele seleciona determinada impressora como padrão

**1. Clone o repositório:**

```bash
git clone https://https://github.com/igorcarmona-git/DefaultPrinter-LDAP.git
cd DefaultPrinter-LDAP
```

**2. Instale as dependências:** (Recomendado via scriptlogon (netlogon))
```bash
python -m pip install --upgrade pip
python -m pip install ldap3 pywin32
```

**3. Exemplo de código de scriptlogon (netlogon)**:
```code
rem Definir variáveis de ambiente com valores específicos
SETX AD_SERVER valueHere
SETX AD_USER valueHere
SETX AD_PASSWORD valueHere
SETX AD_SEARCH_BASE valueHere

rem Pausar para garantir que as variáveis sejam definidas antes de atualizar
timeout /t 5

rem Instalar pacotes Python necessários
echo Instalando pacotes Python necessários
python -m pip install --upgrade pip
python -m pip install ldap3 pywin32

rem Executar o script Python e capturar o código de saída
echo Executando script DefaultPrinter.py.............
set filePrinter=\\serverHere\netlogon\PrinterDefault\PrinterDefault.py
python %filePrinter%

rem Redefinir variáveis de ambiente permanentes para valores iniciais
SETX AD_SERVER 0
SETX AD_USER 0
SETX AD_PASSWORD 0
SETX AD_SEARCH_BASE 0

echo Configurações concluídas com sucesso....[OK]
```
**As variáveis de ambiente definidas no código logon serão usadas no código em momento de execução e depois redefinidas novamente para um valor nulo.**

## Contribuição
Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou pull requests para melhorar o projeto.

**Para mais informações, entrar em contato via redes sociais.**
