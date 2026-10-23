# Projeto Passagem de Turno Inteligente — NOC & Field

MVP operacional em Streamlit para cadastrar pendências de Field Service, consultar/atualizar chamados, registrar a transferência entre turnos e manter uma trilha de auditoria em Excel.

## O que está incluído

- `app.py`: interface Streamlit com abas para novo chamado, consulta/atualização, passagem de turno e histórico de auditoria.
- `src/config.py`: campos, status, prioridades, motivos e turnos padronizados.
- `src/storage.py`: persistência em Excel, controle de gravação via lock, backup diário, IDs de eventos e salvamento atômico.
- `templates/modelo_passagem_turno.xlsx`: modelo Excel com as abas `Chamados`, `Historico`, `Passagens` e `Instrucoes`.
- `data/`: destino do arquivo operacional criado na primeira execução.

## Requisitos

Python 3.10 ou superior recomendado.

## Execução local

No terminal, dentro desta pasta:

```bash
python -m venv .venv
```

Windows PowerShell:

```powershell
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
streamlit run app.py
```

Linux/macOS:

```bash
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

Na primeira inicialização, o sistema copia o modelo para `data/passagem_turno.xlsx`. O arquivo operacional não é o mesmo que o modelo em `templates/`.

## Configurar o caminho do arquivo Excel

Opcionalmente, configure a variável de ambiente `NOC_FIELD_XLSX` para apontar para o arquivo operacional:

PowerShell:

```powershell
$env:NOC_FIELD_XLSX="C:\\NOC\\dados\\passagem_turno.xlsx"
streamlit run app.py
```

Linux/macOS:

```bash
export NOC_FIELD_XLSX="/caminho/seguro/passagem_turno.xlsx"
streamlit run app.py
```

O diretório indicado precisa permitir leitura e gravação pelo processo do Streamlit.

## Fluxo operacional do MVP

1. O operador informa sua identificação na barra lateral.
2. Cadastra um chamado, incluindo loja, prioridade, status, motivo, última atualização e próxima ação.
3. O chamado passa a aparecer na consulta de pendências e pode ser atualizado pelo mesmo ID.
4. Cada alteração relevante gera um evento em `Historico`.
5. Na troca de turno, o operador seleciona os chamados transferidos, origem, destino e resumo. O campo `turno_atual` é atualizado e o evento é registrado no histórico.
6. O sistema cria um backup diário antes da primeira alteração do dia em `data/backups/`.

## Modelo de dados

### Chamados
Um registro por ID de chamado. Contém o estado atual do atendimento, os prazos e os dados da próxima ação.

### Historico
Trilha de auditoria de criação, alterações e passagens. Não apague linhas desta aba.

### Passagens
Uma linha por transferência entre turnos, incluindo os IDs selecionados e o resumo informado.

## Decisões e limites conhecidos

- A identificação digitada pelo operador é apenas informativa e não substitui autenticação corporativa. Antes de produção, integrar login corporativo/SSO ou mecanismo de autenticação aprovado pela empresa.
- Excel é adequado para uma prova de conceito ou piloto controlado. O `FileLock` coordena gravações no mesmo ambiente, mas não transforma um arquivo Excel em banco transacional distribuído. Para vários servidores/instâncias, elevada concorrência ou exigência de disponibilidade, migrar para PostgreSQL ou outra base corporativa.
- Em hospedagens com disco efêmero, o arquivo e os backups locais podem ser perdidos em reinicializações. Use volume persistente/armazenamento corporativo com backup e permissões adequados.
- Evite abrir o arquivo operacional no Excel enquanto o app estiver gravando. Use a interface do Streamlit como método normal de edição.
- O cálculo de SLA deste MVP sinaliza prazos vencidos usando `prazo_sla` e o status atual. Regras de pausa, calendário comercial, feriados e exclusões contratuais ainda precisam ser definidas com a operação antes de usar o indicador oficialmente.
- Os motivos/status estão centralizados em `src/config.py`, para que possam ser ajustados sem procurar valores espalhados na interface.

## Próximos passos recomendados

1. Revisar nomes de status, prioridades, turnos e motivos com o NOC e o Field.
2. Testar com um conjunto controlado de chamados fictícios e validar a passagem entre dois turnos.
3. Validar permissões, backup/recuperação e política de retenção do histórico.
4. Se houver um sistema de chamados com API/exportação, avaliar a integração para evitar redigitação.
5. Só depois consolidar os indicadores gerenciais e regras formais de SLA.

## Observação de integração

O código atual do aplicativo hospedeiro não foi fornecido nesta conversa. Por isso, este pacote é um módulo inicial independente, sem substituir arquivos de um projeto existente. Para incorporar ao app atual, revisar seu `app.py`, a estrutura de páginas, a autenticação e o mecanismo de deploy antes de mesclar.
