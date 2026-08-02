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

## 1. Central Portal account + namespace

1. Sign in at [central.sonatype.com](https://central.sonatype.com/).
2. Open **Namespaces** → create namespace **`ru.myway-line`**.
3. Copy the **Verification Key** shown for the request.
4. At your DNS registrar for `myway-line.ru`, add a **TXT** record:
   - Host / name: `@` (apex of `myway-line.ru`)
   - Value: the Verification Key from the portal
5. Wait until DNS propagates, then verify:

```powershell
Resolve-DnsName myway-line.ru -Type TXT
```

6. Only after the TXT record is visible, confirm verification in the Central Portal.
7. Wait until the namespace status is **Verified**.

Do **not** confirm verification before the TXT record exists — NXDOMAIN can be cached and delay approval.

Official docs: [Register a Namespace](https://central.sonatype.org/register/namespace/).

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
