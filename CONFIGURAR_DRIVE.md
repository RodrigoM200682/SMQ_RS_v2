# Configurar Google Drive — SMQ_RS v3.0

## Como funciona
O app lê o arquivo `.xlsx` diretamente do Google Drive a cada inicialização.
Basta atualizar o arquivo no Drive — na próxima abertura do app, os dados
são atualizados automaticamente.

---

## PASSO 1 — Google Cloud: ativar Drive API e criar conta de serviço

1. Acesse https://console.cloud.google.com
2. Crie um projeto (ex: `SMQ-RS`) ou selecione um existente
3. Menu esquerdo: **APIs e Serviços → Biblioteca**
4. Pesquise **Google Drive API** → **Ativar**
5. Menu esquerdo: **APIs e Serviços → Credenciais**
6. Clique em **+ Criar credenciais → Conta de serviço**
   - Nome: `smq-rs-leitor`
   - Clique em **Criar e continuar → Concluir**
7. Na lista, clique na conta criada → aba **Chaves**
8. **Adicionar chave → Criar nova chave → JSON → Criar**
9. Um arquivo `.json` será baixado — guarde-o

---

## PASSO 2 — Compartilhar a planilha com a conta de serviço

1. Abra https://drive.google.com
2. Localize sua planilha `Consultas_RNC_APP.xlsx`
   (ou faça upload dela se ainda não estiver lá)
3. Clique com botão direito → **Compartilhar**
4. No campo de e-mail, cole o `client_email` do JSON baixado
   (ex: `smq-rs-leitor@smq-rs-123456.iam.gserviceaccount.com`)
5. Permissão: **Leitor** (só precisa de leitura)
6. Clique em **Enviar**

---

## PASSO 3 — Copiar o ID do arquivo

Abra o arquivo no Google Drive e copie o ID da URL:

```
https://drive.google.com/file/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms/view
                               ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
                               Este é o file_id
```

---

## PASSO 4 — Configurar Secrets no Streamlit Cloud

1. Acesse https://share.streamlit.io → seu app → **⋮ → Settings → Secrets**
2. Cole o seguinte (substituindo pelos seus valores):

```toml
[gcp]
file_id = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms"

credentials = '''
{
  "type": "service_account",
  "project_id": "smq-rs-123456",
  "private_key_id": "abc123",
  "private_key": "-----BEGIN RSA PRIVATE KEY-----\nMIIE...\n-----END RSA PRIVATE KEY-----\n",
  "client_email": "smq-rs-leitor@smq-rs-123456.iam.gserviceaccount.com",
  "client_id": "123456789",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/smq-rs-leitor%40smq-rs-123456.iam.gserviceaccount.com"
}
'''
```

3. Clique em **Save** — o app reiniciará

---

## Como usar no dia a dia

1. Quando tiver uma nova planilha:
   - Substitua o arquivo no Google Drive (mesmo nome, mesma pasta)
   - **OU** use o botão **🔄 Atualizar do Drive agora** na sidebar

2. O app carrega a planilha automaticamente a cada abertura

3. Se o Drive estiver offline, usa o cache da última leitura bem-sucedida

---

## Verificar que está funcionando

- Sidebar mostra: **☁️ Google Drive configurado**
- Status mostra: **☁️ Google Drive — dados atualizados**
- Data/hora da última leitura aparece na sidebar
