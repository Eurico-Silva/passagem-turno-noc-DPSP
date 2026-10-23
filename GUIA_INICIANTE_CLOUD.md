# Guia iniciante — publicar o MVP no Streamlit Community Cloud

Este guia parte do princípio de que você tem uma conta Streamlit, mas ainda não tem um aplicativo publicado.

## Antes de começar

1. Crie uma conta no GitHub, se ainda não tiver: https://github.com/signup
2. Entre no Streamlit Community Cloud: https://share.streamlit.io/
3. Conecte/autorizе sua conta GitHub no workspace do Streamlit.

O Community Cloud usa um repositório GitHub para obter o código do aplicativo.

## Criar o repositório do projeto

1. Baixe e extraia `projeto_passagem_turno_mvp.zip` no computador.
2. No GitHub, crie um repositório novo com o nome `passagem-turno-noc-field`.
3. Para um primeiro teste, escolha repositório **público** apenas se não houver informação interna, código confidencial ou dados reais. Não coloque chamados reais no repositório. Se a política da empresa exigir privacidade, use um repositório privado e conceda as permissões necessárias ao Streamlit.
4. Dentro do repositório, envie os arquivos e pastas que estão dentro da pasta `projeto_passagem_turno` extraída. O arquivo `app.py` deve estar na raiz do repositório, ao lado de `requirements.txt` e da pasta `src/`.
5. Confirme se a estrutura ficou semelhante a esta:

```text
passagem-turno-noc-field/
├── app.py
├── requirements.txt
├── README.md
├── src/
│   ├── __init__.py
│   ├── config.py
│   └── storage.py
└── templates/
    └── modelo_passagem_turno.xlsx
```

Não envie a pasta `__pycache__`, a pasta `.venv` nem dados operacionais reais.

## Publicar no Streamlit

1. No workspace do Streamlit Community Cloud, clique em **Create app**.
2. Escolha a opção para implantar um app existente (**Yup, I have an app**).
3. Selecione o repositório `passagem-turno-noc-field`, a branch `main` e o arquivo principal `app.py`.
4. Clique em **Deploy** e acompanhe os logs se houver alguma mensagem de erro.

## Atenção crítica ao armazenamento

Este MVP usa um arquivo Excel salvo no disco local do processo Streamlit. Isso serve para aprender e demonstrar o fluxo com dados fictícios, mas **não é seguro para operar com chamados reais no Streamlit Community Cloud**: a plataforma não garante persistência de arquivos locais; reinicializações ou novas implantações podem apagar o arquivo. Alterações feitas ao Excel no ambiente publicado também não são sincronizadas automaticamente para o GitHub.

Antes do piloto real, precisamos migrar o armazenamento para uma fonte persistente compartilhada (por exemplo, PostgreSQL hospedado ou uma planilha na nuvem com integração adequada) e manter a exportação para Excel. Também será necessário definir autenticação, permissões, backup e recuperação antes de registrar informações operacionais reais.

## Teste inicial

Use somente chamados fictícios. Cadastre um chamado, atualize o status, faça uma passagem entre turnos e confira a aba de histórico. Não use dados de clientes, lojas, técnicos ou incidentes reais até que armazenamento persistente e controle de acesso estejam configurados.
