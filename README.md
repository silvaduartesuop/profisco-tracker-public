# Acompanhamento PROFISCO II — Componente 3

Aplicação web estática (GitHub Pages) para acompanhar o andamento do PROFISCO II e **registrar atualizações em cada campo**. Os dados vêm de `data.json` (gerado a partir da planilha *ACOMPANHAMENTO PROFISCO II*). As edições podem ser **gravadas de volta no próprio repositório** via API do GitHub, usando um token pessoal.

## Conteúdo do repositório

| Arquivo | Função |
|---|---|
| `index.html` | A aplicação (dashboard + edição). Nada mais é necessário para rodar. |
| `data.json` | Os dados do acompanhamento (fonte de verdade). |
| `convert.py` | Converte a planilha `.xlsx` em `data.json` (para atualizar a base a partir do Excel). |
| `.nojekyll` | Faz o GitHub Pages servir os arquivos sem processamento Jekyll. |

## Como publicar no GitHub Pages (uma vez)

1. Crie um repositório no GitHub (ex.: `profisco-tracker`). Pode ser **privado**.
2. Envie estes arquivos para a raiz do repositório (`index.html`, `data.json`, `.nojekyll`, `convert.py`).
   - Pelo site: *Add file → Upload files* → arraste os arquivos → *Commit*.
   - Ou por linha de comando:
     ```bash
     git init && git add . && git commit -m "app de acompanhamento"
     git branch -M main
     git remote add origin https://github.com/SEU_USUARIO/profisco-tracker.git
     git push -u origin main
     ```
3. Em **Settings → Pages**, em *Build and deployment*, selecione **Deploy from a branch**, branch **main** e pasta **/ (root)**. Salve.
4. Aguarde ~1 minuto. O endereço será `https://SEU_USUARIO.github.io/profisco-tracker/`.

> Repositório privado: o Pages exige plano que permita Pages privado. Se preferir, mantenha o repositório privado e rode o `index.html` localmente (basta abrir o arquivo), usando o token para salvar no repo.

## Como salvar atualizações de volta no GitHub (token)

1. No GitHub: **Settings → Developer settings → Personal access tokens → Fine-grained tokens → Generate new token**.
2. Em *Repository access*, escolha **Only select repositories** e selecione este repositório.
3. Em *Permissions → Repository permissions*, defina **Contents: Read and write**.
4. Gere e copie o token (`github_pat_…`).
5. No app, clique em **⚙ Configurar GitHub**, informe *owner*, *repositório*, *branch* (`main`), *caminho* (`data.json`) e cole o **token**. Clique em **Testar conexão** e depois **Salvar configuração**.
6. Ao editar um item e clicar em **Salvar alterações**, os dados ficam salvos no navegador. Clique em **⬆ Salvar no GitHub** para commitar o `data.json` atualizado. Use **⬇ Carregar do GitHub** para trazer a versão mais recente (útil quando várias pessoas editam).

⚠️ **Segurança:** o token fica somente no seu navegador (localStorage). Use sempre um token *fine-grained* limitado a este repositório com escopo mínimo (Contents). Não compartilhe o navegador/perfil com o token salvo.

## Como usar o app

- **Dashboard**: total de itens, concluídos, em andamento, progresso médio e estimativa orçamentária (R$).
- **Filtros**: busca livre + filtro por situação e por área.
- **Tabela**: hierarquia dos produtos (3 → 3.1 → 3.1.1 …). Clique em **Abrir ▸** para editar.
- **Edição por campo**: todos os campos são editáveis. O campo **Progresso (%)** alimenta as barras de andamento. Há um campo **Registrar atualização** que grava uma nota datada no **Histórico** do item (auditoria das mudanças).
- **Exportar/Importar**: baixe `data.json` ou `CSV`, importe um `data.json`, ou recarregue a base original.

## Atualizar a base a partir de uma nova versão da planilha

Quando sair uma nova versão do Excel:

```bash
python convert.py "ACOMPANHAMENTO PROFISCO II - vs XX_XX_XXXX.xlsx"
```

Isso regenera `data.json` (mantendo o mesmo schema). Faça commit do novo `data.json`.
Requer `pip install openpyxl`.

## Observações

- O **Progresso (%)** não existia na planilha; foi criado como campo editável, com valor inicial derivado da *Situação* (Concluído = 100, Em Andamento = 40, Não iniciado = 0). Ajuste manualmente conforme o real.
- Datas inconsistentes na planilha original (ex.: um valor `23042023`) são mantidas como texto para você corrigir no app.
