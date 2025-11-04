# 🤖 Automação de Redirecionamentos Tray (Python + Selenium)

Automação em **Python** para aplicar redirecionamentos 301 no painel administrativo da **Tray Commerce**.  
Ideal para quem precisa importar grandes volumes de redirecionamentos **DE → PARA** sem depender de processos manuais demorados.

---

## ⚡ Objetivo

Reduzir o tempo e custo de aplicação de redirecionamentos em massa, automatizando a inserção das URLs via Selenium e controlando o progresso em uma planilha Excel.

---

## 🚀 Funcionalidades

- **Leitura de Planilha (.xlsx)** com as colunas `DE`, `PARA` e `Redirect?`
- **Ignora linhas processadas** automaticamente
- **Controle via Chrome WebDriver** com Selenium
- **Timeout inteligente (20s)** para login manual na Tray
- **Atualiza status TRUE** após sucesso em cada redirecionamento
- **Execução segura**, com espera explícita e refresh automático em caso de erro

---

## 🧩 Estrutura da Planilha

| DE (Coluna A) | PARA (Coluna B) | Redirect? (Coluna C) |
|----------------|----------------|-----------------------|
| /produtos/url-antiga | /colecao/url-nova | TRUE ou (vazio) |

**Importante:** Use apenas o caminho das URLs (sem domínio).  
Exemplo correto: `/produtos/garrafa-termica`  
Exemplo incorreto: `https://minhaloja.com.br/produtos/garrafa-termica`