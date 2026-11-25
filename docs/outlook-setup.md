# Registrare Sweep su Microsoft Entra (Outlook / Hotmail)

Outlook.com non accetta più password da IMAP/SMTP: Sweep entra con il login Microsoft (OAuth2).
Per farlo serve un **client ID**, ottenuto registrando l'app una sola volta. Vale per tutti gli utenti di `Sweep.exe`.

Il client ID **non è un segreto**: identifica l'app, non l'utente, e un'app desktop non ha password propria.
Per questo si scrive direttamente nel codice (`MS_CLIENT_ID` in `app.py`).

Serve un account Microsoft che possa registrare app: un account di lavoro/scuola (se l'IT non lo blocca)
o un tenant Azure gratuito. Un semplice account Hotmail senza tenant non basta.

## 1. Crea la registrazione

1. Vai su [entra.microsoft.com](https://entra.microsoft.com) e accedi.
2. Menu a sinistra: **Identity → Applications → App registrations** (in alcune versioni **Entra ID → App registrations**).
3. **+ New registration**:
   - **Name**: `Sweep`
   - **Supported account types**: **Personal Microsoft accounts only** ← fondamentale: Sweep usa l'authority `/consumers`.
   - **Redirect URI**: lascia vuoto.
4. **Register**.

## 2. Aggiungi il redirect

1. Nell'app → **Authentication** → **+ Add a platform**.
2. Scegli **Mobile and desktop applications**.
3. In **Custom redirect URIs** scrivi esattamente `http://localhost` (niente porta, niente `/` finale).
4. **Configure**.
5. Sempre in *Authentication*, in fondo: **Allow public client flows** → **Yes** → **Save**.

## 3. Permessi (facoltativo)

Sweep chiede i permessi da solo al login (*dynamic consent*). Se vuoi vederli elencati:

1. **API permissions → + Add a permission → Microsoft Graph → Delegated permissions**.
2. Spunta `IMAP.AccessAsUser.All` e `SMTP.Send` → **Add permissions**.

Niente "Grant admin consent": con gli account personali ogni utente acconsente da sé al primo login.

## 4. Copia l'ID

In **Overview** copia **Application (client) ID** (formato `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`).
**Non** creare un client secret.

## 5. Prova

Per provarlo senza toccare il codice:

```bash
SWEEP_MS_CLIENT_ID=<il-tuo-id> python app.py
```

In Sweep: provider **Outlook / Hotmail**, inserisci l'indirizzo, **Sign in with Microsoft**.
Si apre il browser: accedi e accetta i permessi. L'account compare nella lista, poi **Scan inboxes**.

Se funziona, scrivi l'ID in `MS_CLIENT_ID` in `app.py` e ricompila l'exe.

## Problemi comuni

| Errore | Causa |
|---|---|
| `AADSTS50020` / `unauthorized_client` | Tipo di account diverso da *Personal Microsoft accounts only* (passo 1). |
| `redirect_uri mismatch` | Redirect non esattamente `http://localhost` o piattaforma diversa da *Mobile and desktop* (passo 2). |
| "AUTHENTICATE failed" dopo un login riuscito | IMAP spento: Outlook.com → Impostazioni → Posta → Inoltro e IMAP → attivalo. |
| "Microsoft sign-in did not finish" | Scheda del browser chiusa o login non completato entro 5 minuti: riprova. |
