# Define o caminho da pasta temporária onde os arquivos da instalação serão armazenados
$caminhoPastaTemp = "$env:TEMP\RicohInstaller"

# URL oficial do driver da impressora Ricoh para download
$urldoDriveRicoh = "https://support.ricoh.com/bb/pub_e/dr_ut_e/0001333/0001333436/V108/z97664L15.exe"

# URL do arquivo de resposta para instalação silenciosa (setup.iss), hospedado no Google Drive
$baixarInstallDriveRicoh = "https://drive.google.com/uc?export=download&id=1cBP4UdKU5BWfBfgCck-ZSoFtSTE5dX-T"

# Caminho completo onde o executável do driver será salvo
$exeDrive = "$caminhoPastaTemp\ricoh_driver.exe"

# Caminho completo para salvar o arquivo de resposta da instalação
$arquivoRespostaInstall = "$caminhoPastaTemp\setup.iss"

# Caminho para o arquivo de log da instalação
$caminhoLogInstall = "$caminhoPastaTemp\install_log.txt"

# Solicita ao usuário o IP da impressora
$ipMaquina = Read-Host "Digite o IP da impressora (ex: 192.168.0.100)"

# Solicita ao usuário um nome para identificar a impressora
$nomeMaquina = Read-Host "Digite um nome de identificação da impressora: "

# Define o nome da impressora com base na entrada do usuário
$printerName = $nomeMaquina

# Cria um nome de porta no formato IP_<endereço IP>
$nomePorta = "IP_$ipMaquina"

# Define o nome do driver da impressora que será utilizado
$nomeDriver = "RICOH SP 3710SF PCL 6"

# Verifica se a pasta temporária já existe; se não, cria a pasta
if (!(Test-Path $caminhoPastaTemp)) {
    New-Item -Path $caminhoPastaTemp -ItemType Directory | Out-Null
}

# Informa o início do download do driver
Write-Host "Baixando driver..."
# Baixa o executável do driver da URL fornecida e salva no caminho definido
Invoke-WebRequest -Uri $urldoDriveRicoh -OutFile $exeDrive

# Informa o início do download do arquivo de resposta
Write-Host "Baixando arquivo de resposta..."
# Baixa o arquivo de resposta da URL e salva no caminho definido
Invoke-WebRequest -Uri $baixarInstallDriveRicoh -OutFile $arquivoRespostaInstall

# Informa o início da instalação do driver
Write-Host "Instalando driver..."
# Executa a instalação do driver em modo silencioso usando o arquivo de resposta e gera um log
Start-Process -FilePath $exeDrive -ArgumentList "/s", "/f1=$arquivoRespostaInstall", "/f2=$caminhoLogInstall" -Wait
# Informa que a instalação foi concluída
Write-Host "Driver instalado com sucesso.`n"

# Verifica se a porta da impressora já existe
if (!(Get-PrinterPort -Name $nomePorta -ErrorAction SilentlyContinue)) {
    # Se a porta não existir, cria uma nova porta TCP/IP com o IP informado
    Add-PrinterPort -Name $nomePorta -PrinterHostAddress $ipMaquina
    Write-Host "Porta criada: $nomePorta"
} else {
    # Caso a porta já exista, informa o usuário
    Write-Host "Porta já existente: $nomePorta"
}

# Adiciona a impressora usando o nome, porta e driver definidos
Add-Printer -Name $printerName -PortName $nomePorta -DriverName $nomeDriver -ErrorAction Stop
# Informa que a impressora foi instalada com sucesso
Write-Host "Impressora '$printerName' instalada com sucesso!"
