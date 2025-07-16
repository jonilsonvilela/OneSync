# OneSync - JOCA

**OneSync - JOCA** é um sistema de automação com interface web projetado para otimizar e acelerar a atualização de dados de processos jurídicos no sistema Legal One, utilizando uma planilha Excel como entrada.

A ferramenta foi desenvolvida para equipes jurídicas que buscam mais controle, agilidade e rastreabilidade no gerenciamento de grandes volumes de processos, eliminando erros manuais e liberando os profissionais para atividades de maior valor estratégico.

## Funcionalidades Principais

* **Interface Web Intuitiva:** Uma interface moderna e responsiva para upload de planilhas e acompanhamento da execução.
* **Automação com RPA:** Um robô construído com Playwright que realiza o login, busca e preenchimento dos dados no sistema Legal One.
* **Dashboard de Análise:** Métricas e gráficos sobre a última execução, incluindo total de processos, taxas de sucesso/falha e tempo gasto por etapa.
* **Histórico de Execuções:** Todas as execuções ficam salvas e podem ser consultadas ou baixadas a qualquer momento.
* **Modo Noturno:** Um seletor de tema para uma experiência de visualização mais confortável.
* **Log em Tempo Real:** Acompanhe o progresso do robô em tempo real diretamente pela interface web.

## 🚀 Tecnologias Utilizadas

* **Backend:** Python 3.10+ com Flask
* **Frontend:** HTML5, Tailwind CSS, Chart.js
* **Automação RPA:** Playwright
* **Manipulação de Dados:** Pandas

## ⚙️ Instalação e Configuração

Siga os passos abaixo para configurar o ambiente e executar o projeto localmente.

**1. Clone o Repositório**
```bash
git clone <url-do-seu-repositorio>
cd OneSync-JOCA
```

**2. Crie e Ative um Ambiente Virtual**
```bash
# Windows
python -m venv venv
.\venv\Scripts\activate

# macOS / Linux
python3 -m venv venv
source venv/bin/activate
```

**3. Instale as Dependências**
```bash
pip install -r requirements.txt
```

**4. Instale os Navegadores do Playwright**
```bash
playwright install
```

**5. Configure as Variáveis de Ambiente**
O robô precisa das suas credenciais do Legal One para funcionar. Crie um arquivo `.env` na raiz do projeto ou configure as variáveis de ambiente no seu sistema com os seguintes nomes:
```
LEGALONE_USUARIO="seu_usuario_aqui"
LEGALONE_SENHA="sua_senha_aqui"
```
*O `app.py` precisa ser ajustado para ler essas variáveis de um arquivo `.env` se você escolher essa abordagem (ex: usando a biblioteca `python-dotenv`).*

**6. Execute a Aplicação**
```bash
python app.py
```
A aplicação estará disponível em `http://127.0.0.1:8080`.

## (Como Usar)

1.  Acesse a aplicação pelo navegador.
2.  Na página **Dashboard**, faça o upload da sua planilha `entrada.xlsx`.
3.  Clique no botão **"▶️ Executar Atualização"**.
4.  Você será redirecionado para a tela de execução, onde poderá acompanhar o log em tempo real.
5.  Após a finalização, acesse a aba **"Resultado"** para ver o dashboard analítico ou a aba **"Histórico"** para consultar execuções passadas.

---