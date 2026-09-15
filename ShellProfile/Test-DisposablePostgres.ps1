# Isolated Docker smoke test: no host port, host mounts, production credentials or network.
[CmdletBinding()]
param([string]$Image = 'postgres:17.11')
$ErrorActionPreference = 'Stop'
$name = 'gsadus-postgres-test-' + [guid]::NewGuid().ToString('N').Substring(0,12)
$created = $false
try {
    docker --context desktop-linux info --format '{{.ServerVersion}}'
    if ($LASTEXITCODE -ne 0) { throw 'Docker Desktop Linux engine is not ready.' }
    docker --context desktop-linux run --detach --name $name --network none `
        --label gsadus.purpose=disposable-test --tmpfs /var/lib/postgresql/data `
        --env POSTGRES_HOST_AUTH_METHOD=trust $Image | Out-Null
    if ($LASTEXITCODE -ne 0) { throw 'Could not start disposable PostgreSQL.' }
    $created = $true
    $ready = $false
    for ($attempt = 0; $attempt -lt 30; $attempt++) {
        docker --context desktop-linux exec $name pg_isready -U postgres *> $null
        if ($LASTEXITCODE -eq 0) { $ready = $true; break }
        Start-Sleep -Seconds 1
    }
    if (-not $ready) { throw 'PostgreSQL did not become ready.' }
    @'
BEGIN;
CREATE TABLE isolated_check (id integer PRIMARY KEY, label text NOT NULL);
INSERT INTO isolated_check VALUES (1, 'disposable');
DO $$ BEGIN
  IF (SELECT count(*) FROM isolated_check) <> 1 THEN RAISE EXCEPTION 'insert failed'; END IF;
END $$;
ROLLBACK;
DO $$ BEGIN
  IF to_regclass('public.isolated_check') IS NOT NULL THEN RAISE EXCEPTION 'rollback failed'; END IF;
END $$;
SELECT 'isolated SQL test passed' AS result;
'@ | docker --context desktop-linux exec --interactive $name psql -U postgres -v ON_ERROR_STOP=1
    if ($LASTEXITCODE -ne 0) { throw 'Isolated SQL test failed.' }
    docker --context desktop-linux stop --time 10 $name | Out-Null
    if ($LASTEXITCODE -ne 0) { throw 'Container stop failed.' }
    $code = docker --context desktop-linux inspect --format '{{.State.ExitCode}}' $name
    if ($LASTEXITCODE -ne 0 -or $code -ne '0') { throw 'PostgreSQL did not stop cleanly.' }
    Write-Host 'PASS: Docker PostgreSQL started, SQL transaction and rollback passed, graceful shutdown exit 0.'
} finally {
    if ($created) {
        docker --context desktop-linux rm --force --volumes $name | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "Could not remove test container $name" }
        Write-Host "Removed test container: $name"
    }
}
