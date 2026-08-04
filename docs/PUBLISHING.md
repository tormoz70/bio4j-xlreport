# Publishing to Maven Central

Coordinates use the reverse DNS of **myway-line.ru**:

| Field | Value |
| --- | --- |
| Namespace (register) | `ru.myway-line` |
| groupId | `ru.myway-line.xlreport` |
| Artifacts | `engine-core`, `engine-poi`, `engine-js` |
| Example | `ru.myway-line.xlreport:engine-core:0.1.0` |

`app-console` is an application runner and is **not** published to Maven Central.

Java packages: `ru.mywayline.xlreport.*` (hyphen removed — Java identifiers cannot contain `-`).  
Maven `groupId` keeps the domain form: `ru.myway-line.xlreport`.

## Checklist

| Step | Status | Where |
| --- | --- | --- |
| 1. Gradle publish config | Done in repo | `build.gradle` |
| 2. Central Portal account | Manual | [central.sonatype.com](https://central.sonatype.com/) |
| 3. Namespace `ru.myway-line` + DNS TXT | Manual | Domain registrar for `myway-line.ru` |
| 4. User token | Manual | Central Portal → Account → User Token |
| 5. GPG key + publish to keyserver | Manual | Gpg4win / GnuPG |
| 6. `~/.gradle/gradle.properties` secrets | Manual | Local machine only |
| 7. `gradlew publishAndReleaseToMavenCentral` | After 2–6 | Terminal |

## 1. Central Portal account + namespace (detailed)

Official docs: [Create account](https://central.sonatype.org/register/central-portal/), [Register namespace](https://central.sonatype.org/register/namespace/).

### 1.1 Create / sign in to Central Portal

1. Open https://central.sonatype.com/
2. Click **Sign In** (top right).
3. Sign up with **GitHub**, **Google**, or email/password.
   - Prefer the same GitHub account as `tormoz70` if possible.
   - Use a real email you can access (needed for account recovery/support).
4. Complete signup (confirm email if asked).
5. You should land on the Central Publisher Portal dashboard.

Note: if you signed in with GitHub, Sonatype may auto-create `io.github.tormoz70`. That is fine — you still need a **separate** DNS namespace `ru.myway-line` for this project.

### 1.2 Request namespace `ru.myway-line`

Why this value:

| Domain you own | Namespace to register |
| --- | --- |
| `myway-line.ru` | `ru.myway-line` |

After verification you can publish any groupId under it, including our project group:

`ru.myway-line.xlreport`

Steps:

1. Top-right corner → click your **username/email**.
2. Click **View Namespaces**.
3. Click **Add Namespace**.
4. Enter exactly:

```text
ru.myway-line
```

5. Click **Submit**.
6. The new namespace appears with status **Unverified**.
7. On the namespace card, find **Verification Key** and click the clipboard icon to copy it.
   - It looks like a token / key string (sometimes similar to older `OSSRH-…` style ids).
   - Keep this key — it goes into DNS in step 2.

Do **not** click **Verify Namespace / Confirm** yet.

---

## 2. DNS TXT verification for `myway-line.ru` (detailed)

Official FAQ: [How to set TXT record](https://central.sonatype.org/faq/how-to-set-txt-record/).

Sonatype checks the **apex domain** matching the namespace:

| Namespace | DNS host they query |
| --- | --- |
| `ru.myway-line` | `myway-line.ru` |

They do **not** check `www.myway-line.ru`, `maven.myway-line.ru`, etc.

### 2.1 Add TXT record at your registrar

1. Log in to the panel where DNS for `myway-line.ru` is managed
   (registrar / hosting / Cloudflare / etc.).
2. Open DNS zone for **`myway-line.ru`**.
3. Add a new record:

| Field | What to set |
| --- | --- |
| Type | `TXT` |
| Host / Name | `@` or blank or `myway-line.ru` (depends on panel; must apply to apex) |
| Value / Content | paste the **Verification Key** from Central Portal (exact, no quotes unless the panel requires them) |
| TTL | default or `3600` |

Common panel gotchas:

- If Host is `@` → record applies to `myway-line.ru` (correct).
- If you type Host `myway-line.ru` in a panel that already suffixes the domain, you may accidentally create `myway-line.ru.myway-line.ru` — wrong.
- Do not put the key on a subdomain unless Sonatype asked for that (they didn't).
- Existing other TXT records (SPF, DKIM, etc.) can stay; add one more TXT.

### 2.2 Wait for DNS propagation, then check locally

In PowerShell:

```powershell
Resolve-DnsName myway-line.ru -Type TXT
```

Or CMD:

```bat
nslookup -type=TXT myway-line.ru
```

Online check: https://toolbox.googleapps.com/apps/dig/#TXT/ → query `myway-line.ru`

You must see your Verification Key among the TXT answers.

If not visible yet:

- wait 5–30 minutes (sometimes up to a few hours),
- confirm you edited the **authoritative** DNS (not an unused secondary zone),
- flush local cache if needed: `ipconfig /flushdns`.

### 2.3 Only then confirm verification in Central Portal

1. Return to https://central.sonatype.com/ → **View Namespaces**.
2. On `ru.myway-line` click **Verify Namespace**.
3. Confirm only after TXT is visible in `Resolve-DnsName`.
4. Status becomes **Verification Pending**, then usually **Verified** within minutes.

Critical: if you confirm **before** TXT exists, Sonatype may cache NXDOMAIN and verification can stall for hours. Use **Cancel Verification**, fix DNS, then verify again.

### 2.4 Success criteria

Namespace `ru.myway-line` status = **Verified**.

Then you may publish coordinates like:

```text
ru.myway-line.xlreport:engine-core:0.1.0
```

## 2. User token

In Central Portal → account → **Generate User Token**.  
You get a username/password pair used only for publishing (not your login password).

## 3. GPG signing key (Windows)

Maven Central requires signed artifacts. On this machine install GnuPG first, for example:

```powershell
winget install --id GnuPG.GnuPG -e
```

Then open a **new** terminal and:

```powershell
gpg --full-generate-key
# RSA and RSA, 4096 bits, no expiry (or set one), Real name + email, passphrase

gpg --list-secret-keys --keyid-format LONG
# Note the KEY_ID after rsa4096/ (e.g. ABCDEF1234567890)

gpg --keyserver hkps://keys.openpgp.org --send-keys <KEY_ID>
# Fallback: gpg --keyserver keyserver.ubuntu.com --send-keys <KEY_ID>
```

Optional export for in-memory Gradle signing:

```powershell
gpg --export-secret-keys --armor <KEY_ID>
```

## 4. Local Gradle secrets

Create `%USERPROFILE%\.gradle\gradle.properties` (never commit this file):

```properties
mavenCentralUsername=<token username>
mavenCentralPassword=<token password>

signing.keyId=<last 8 hex of key id>
signing.password=<gpg passphrase>
signing.secretKeyRingFile=<path to secring.gpg>
```

Modern GnuPG often has no `secring.gpg`. Prefer in-memory key instead:

```properties
mavenCentralUsername=<token username>
mavenCentralPassword=<token password>
signingInMemoryKey=<paste ASCII-armored private key; keep newlines as \n or use a single-line escaped form>
signingInMemoryKeyId=<KEY_ID>
signingInMemoryKeyPassword=<gpg passphrase>
```

Signing is enabled automatically when `signing.keyId` or `signingInMemoryKey` is present. Without them, `publishToMavenLocal` still works; Maven Central upload will fail validation unsigned.

## 5. Publish

Dry-run locally (Maven Local):

```powershell
.\gradlew.bat publishToMavenLocal
```

Upload to Central Portal (manual Publish in UI):

```powershell
.\gradlew.bat publishToMavenCentral
```

Then open [Deployments](https://central.sonatype.com/) → **Publish**.

Or upload and release in one step:

```powershell
.\gradlew.bat publishAndReleaseToMavenCentral
```

Sync to `repo1.maven.org` usually takes 10–30 minutes after the deployment reaches **Published**.

## 6. Consume

```gradle
dependencies {
    implementation 'ru.myway-line.xlreport:engine-core:0.1.0'
    implementation 'ru.myway-line.xlreport:engine-poi:0.1.0'
    implementation 'ru.myway-line.xlreport:engine-js:0.1.0'
}
```
