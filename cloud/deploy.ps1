<#
Qent -> Azure (EU). Readable `az` steps, run top to bottom after `az login`. Nothing here has been run for you.
Review the variables, then run each block. Costs: PostgreSQL B1ms + Container Apps (1 replica) + Azure OpenAI (pay per token).
Not covered: custom domain, WAF, multi-replica. # ponytail: move to Bicep if you need repeatable environments.
#>
$ErrorActionPreference = "Stop"

# ---- 1. names and region (EU only; confirm with your data policy) -------------------------------------
$rg      = "rg-qent-eu"
$loc     = "swedencentral"          # or northeurope
$suffix  = -join ((97..122) | Get-Random -Count 5 | ForEach-Object { [char]$_ })
$env_    = "cae-qent"
$app     = "qent-web"
$job     = "qent-collect"
$acr     = "acrqent$suffix"
$kv      = "kv-qent-$suffix"
$pg      = "pg-qent-$suffix"
$aoai    = "aoai-qent-$suffix"
$logws   = "log-qent"
$img     = "qent:1"

az group create -n $rg -l $loc | Out-Null
az monitor log-analytics workspace create -g $rg -n $logws -l $loc | Out-Null
$logId  = az monitor log-analytics workspace show -g $rg -n $logws --query customerId -o tsv
$logKey = az monitor log-analytics workspace get-shared-keys -g $rg -n $logws --query primarySharedKey -o tsv

# ---- 2. PostgreSQL (private to Azure services; firewall rule 0.0.0.0 = "Azure services only", not the internet) ----
$pgPass = -join ((48..57) + (65..90) + (97..122) | Get-Random -Count 32 | ForEach-Object { [char]$_ })
az postgres flexible-server create -g $rg -n $pg -l $loc --tier Burstable --sku-name Standard_B1ms --storage-size 32 `
    --version 16 --admin-user qentadmin --admin-password $pgPass --database-name qent --public-access 0.0.0.0 --backup-retention 7 | Out-Null
$dbUrl = "postgresql://qentadmin:$pgPass@$pg.postgres.database.azure.com:5432/qent?sslmode=require"

# ---- 3. Azure OpenAI in an EU Data Zone (deployment type DataZoneStandard keeps processing in the EU) ------
az cognitiveservices account create -g $rg -n $aoai -l $loc --kind OpenAI --sku S0 --custom-domain $aoai | Out-Null
az cognitiveservices account deployment create -g $rg -n $aoai --deployment-name qent-chat `
    --model-name gpt-4.1-mini --model-version 2025-04-14 --model-format OpenAI --sku-name DataZoneStandard --sku-capacity 20 | Out-Null
# if that model/version is unavailable in your region: `az cognitiveservices model list -l $loc` and pick a small chat model
$aoaiEndpoint = az cognitiveservices account show -g $rg -n $aoai --query properties.endpoint -o tsv
$aoaiKey      = az cognitiveservices account keys list -g $rg -n $aoai --query key1 -o tsv

# ---- 4. Key Vault holds the secrets; the apps read them as secret references -----------------------------------
az keyvault create -g $rg -n $kv -l $loc --enable-rbac-authorization true | Out-Null
$me = az ad signed-in-user show --query id -o tsv
az role assignment create --assignee $me --role "Key Vault Secrets Officer" --scope (az keyvault show -n $kv --query id -o tsv) | Out-Null
az keyvault secret set --vault-name $kv -n database-url --value $dbUrl | Out-Null
az keyvault secret set --vault-name $kv -n azure-openai-key --value $aoaiKey | Out-Null

# ---- 5. Build the image in the registry (no local Docker needed) ---------------------------------------------------
az acr create -g $rg -n $acr --sku Basic --admin-enabled false | Out-Null
az acr build -r $acr -t $img -f cloud/Dockerfile . | Out-Null

# ---- 6. Container Apps environment, identity, web app ---------------------------------------------------------------------
az containerapp env create -g $rg -n $env_ -l $loc --logs-workspace-id $logId --logs-workspace-key $logKey | Out-Null
$uami = az identity create -g $rg -n id-qent --query id -o tsv
$uamiPrincipal = az identity show --ids $uami --query principalId -o tsv
az role assignment create --assignee-object-id $uamiPrincipal --assignee-principal-type ServicePrincipal --role AcrPull --scope (az acr show -n $acr --query id -o tsv) | Out-Null
az role assignment create --assignee-object-id $uamiPrincipal --assignee-principal-type ServicePrincipal --role "Key Vault Secrets User" --scope (az keyvault show -n $kv --query id -o tsv) | Out-Null
Start-Sleep 30   # role assignments need a moment to propagate

$secrets = "database-url=keyvaultref:https://$kv.vault.azure.net/secrets/database-url,identityref:$uami", "azure-openai-key=keyvaultref:https://$kv.vault.azure.net/secrets/azure-openai-key,identityref:$uami"
$envVars = "DATABASE_URL=secretref:database-url", "AZURE_OPENAI_KEY=secretref:azure-openai-key", "AZURE_OPENAI_ENDPOINT=$aoaiEndpoint", "AZURE_OPENAI_DEPLOYMENT=qent-chat", "LLM_PROVIDER=azure", "REQUIRE_AUTH=1"

az containerapp create -g $rg -n $app --environment $env_ --image "$acr.azurecr.io/$img" --registry-server "$acr.azurecr.io" --registry-identity $uami `
    --user-assigned $uami --target-port 8080 --ingress external --min-replicas 1 --max-replicas 1 --cpu 0.5 --memory 1Gi `
    --secrets $secrets --env-vars $envVars | Out-Null
$fqdn = az containerapp show -g $rg -n $app --query properties.configuration.ingress.fqdn -o tsv
az containerapp update -g $rg -n $app --set-env-vars "ALLOWED_HOSTS=$fqdn" | Out-Null   # Host allowlist (DNS rebinding guard)
az containerapp ingress sticky-sessions set -g $rg -n $app --affinity none 2>$null

# ---- 7. Entra ID sign-in (Easy Auth) with Viewer / Analyst / Admin app roles ---------------------------------------------
$clientId = az ad app create --display-name "Qent" --sign-in-audience AzureADMyOrg `
    --web-redirect-uris "https://$fqdn/.auth/login/aad/callback" --enable-id-token-issuance true --query appId -o tsv
$roles = @(
    @{ allowedMemberTypes = @("User"); description = "Read the dashboard";                     displayName = "Viewer";  id = [guid]::NewGuid(); isEnabled = $true; value = "Viewer" },
    @{ allowedMemberTypes = @("User"); description = "Mark items and set credibility";          displayName = "Analyst"; id = [guid]::NewGuid(); isEnabled = $true; value = "Analyst" },
    @{ allowedMemberTypes = @("User"); description = "Also grade sources and run collection";   displayName = "Admin";   id = [guid]::NewGuid(); isEnabled = $true; value = "Admin" }
) | ConvertTo-Json -Depth 4
$rolesFile = New-TemporaryFile; Set-Content $rolesFile $roles
az ad app update --id $clientId --app-roles "@$rolesFile"; Remove-Item $rolesFile
az ad sp create --id $clientId | Out-Null
# Only assigned users can sign in. Assign people to the Viewer/Analyst/Admin roles on the "Qent" enterprise app in the portal.
$spId = az ad sp show --id $clientId --query id -o tsv
az ad sp update --id $spId --set appRoleAssignmentRequired=true

$tenant = az account show --query tenantId -o tsv
az containerapp auth microsoft update -g $rg -n $app --client-id $clientId --issuer "https://login.microsoftonline.com/$tenant/v2.0" --yes | Out-Null
az containerapp auth update -g $rg -n $app --unauthenticated-client-action RedirectToLoginPage --redirect-provider azureactivedirectory --excluded-paths "/healthz" | Out-Null
# Google as a second provider later: `az containerapp auth google update ...` (no code change in Qent).

# ---- 7b. MISP (e.g. FinMISP): keys never go in the Qent UI. In the UI add a source of kind "MISP instance" and type the NAME of an
#          environment variable. Here you create that variable from a Key Vault secret, on BOTH the web app and the Job:
# az keyvault secret set --vault-name $kv -n finmisp-key --value "<the MISP API key>"
# $secrets += "finmisp-key=keyvaultref:https://$kv.vault.azure.net/secrets/finmisp-key,identityref:$uami"
# az containerapp secret set -g $rg -n $app --secrets "finmisp-key=keyvaultref:https://$kv.vault.azure.net/secrets/finmisp-key,identityref:$uami"
# az containerapp update -g $rg -n $app --set-env-vars "FINMISP_KEY=secretref:finmisp-key"     # same two commands for the Job: az containerapp job ...
# A MISP on a private network is refused by the SSRF guard. Allow that one host at deploy time (never from the UI):
# az containerapp update -g $rg -n $app --set-env-vars "ALLOW_PRIVATE_HOSTS=misp.internal.example"   # https only, certificate must validate

# ---- 8. Scheduled collection: same image, every 6 hours ------------------------------------------------------------------------
az containerapp job create -g $rg -n $job --environment $env_ --trigger-type Schedule --cron-expression "0 */6 * * *" `
    --image "$acr.azurecr.io/$img" --registry-server "$acr.azurecr.io" --registry-identity $uami --mi-user-assigned $uami `
    --cpu 0.5 --memory 1Gi --replica-timeout 900 --command "python" "qent.py" "collect" `
    --secrets $secrets --env-vars "DATABASE_URL=secretref:database-url", "PORT=8080" | Out-Null

# ---- 9. First data: import the old workbooks once (run from this folder, needs `pip install openpyxl "psycopg[binary]"`) ----------
# $env:DATABASE_URL = $dbUrl; python qent.py import pir_workbook.xlsx collected_auto.xlsx
# (add your client IP to the Postgres firewall first: az postgres flexible-server firewall-rule create ... then remove it again)

# ---- 10. Egress allowlist (SSRF backstop): put the environment in a VNet with an Azure Firewall / NSG that allows only
#          the feed hosts in FEEDS, article hosts you trust, the Azure OpenAI endpoint and Key Vault. Not scripted here. ----------
Write-Host "Qent: https://$fqdn   (assign users to the Viewer/Analyst/Admin roles on the 'Qent' enterprise app)"
Write-Host "Rotate any API key that was ever pasted into chat or a file."
