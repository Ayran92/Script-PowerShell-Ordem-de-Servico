Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# === CONFIGURAÇÕES ===
$diretorioHistorico = "C:\Users\Vava\Desktop\Teste Codigo"
$nomeImpressora = "MP-4200 TH"   # Nome exato conforme "Get-Printer"

# === Função de Impressão RAW Corrigida ===
function EnviarParaImpressoraRaw($printerName, $text) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class RawPrinterHelper {
    [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Ansi)]
    public class DOCINFOA {
        [MarshalAs(UnmanagedType.LPStr)] public string pDocName;
        [MarshalAs(UnmanagedType.LPStr)] public string pOutputFile;
        [MarshalAs(UnmanagedType.LPStr)] public string pDataType;
    }
    [DllImport("winspool.Drv", EntryPoint="OpenPrinterA", SetLastError=true)]
    public static extern bool OpenPrinter(string szPrinter, out IntPtr hPrinter, IntPtr pd);
    [DllImport("winspool.Drv", EntryPoint="ClosePrinter", SetLastError=true)]
    public static extern bool ClosePrinter(IntPtr hPrinter);
    [DllImport("winspool.Drv", EntryPoint="StartDocPrinterA", SetLastError=true)]
    public static extern bool StartDocPrinter(IntPtr hPrinter, int Level, [In, MarshalAs(UnmanagedType.LPStruct)] DOCINFOA di);
    [DllImport("winspool.Drv", EntryPoint="StartPagePrinter", SetLastError=true)]
    public static extern bool StartPagePrinter(IntPtr hPrinter);
    [DllImport("winspool.Drv", EntryPoint="EndPagePrinter", SetLastError=true)]
    public static extern bool EndPagePrinter(IntPtr hPrinter);
    [DllImport("winspool.Drv", EntryPoint="EndDocPrinter", SetLastError=true)]
    public static extern bool EndDocPrinter(IntPtr hPrinter);
    [DllImport("winspool.Drv", EntryPoint="WritePrinter", SetLastError=true)]
    public static extern bool WritePrinter(IntPtr hPrinter, byte[] pBytes, int dwCount, out int dwWritten);
}
"@

    # Usa codificação compatível com impressoras térmicas
    $bytes = [System.Text.Encoding]::GetEncoding(850).GetBytes($text + "`r`n`r`n`r`n")

    $docinfo = New-Object RawPrinterHelper+DOCINFOA
    $docinfo.pDocName = "OrdemServico_VavaCell"
    $docinfo.pDataType = "RAW"

    [IntPtr]$hPrinter = [IntPtr]::Zero
    if ([RawPrinterHelper]::OpenPrinter($printerName, [ref]$hPrinter, [IntPtr]::Zero)) {
        [RawPrinterHelper]::StartDocPrinter($hPrinter, 1, $docinfo) | Out-Null
        [RawPrinterHelper]::StartPagePrinter($hPrinter) | Out-Null
        [int]$written = 0
        [RawPrinterHelper]::WritePrinter($hPrinter, $bytes, $bytes.Length, [ref]$written) | Out-Null
        [RawPrinterHelper]::EndPagePrinter($hPrinter) | Out-Null
        [RawPrinterHelper]::EndDocPrinter($hPrinter) | Out-Null
        [RawPrinterHelper]::ClosePrinter($hPrinter) | Out-Null
    } else {
        throw "Não foi possível abrir a impressora $printerName"
    }
}

# === CRIA A JANELA ===
$form = New-Object System.Windows.Forms.Form
$form.Text = "Cadastro de Ordem de Serviço - VAVACELL"
$form.Size = New-Object System.Drawing.Size(400,480)
$form.StartPosition = "CenterScreen"

# Função para criar rótulos e caixas de texto
function NovaEntrada($labelText, [ref]$yPos) {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $labelText
    $label.Location = New-Object System.Drawing.Point(20, $yPos.Value)
    $label.AutoSize = $true
    $form.Controls.Add($label)

    $yPos.Value += 20
    $textbox = New-Object System.Windows.Forms.TextBox
    $textbox.Location = New-Object System.Drawing.Point(20, $yPos.Value)
    $textbox.Width = 330
    $form.Controls.Add($textbox)

    $yPos.Value += 40
    return $textbox
}

# === CAMPOS ===
$y = 20
$txtNome        = NovaEntrada "Nome do cliente:" ([ref]$y)
$txtSenha       = NovaEntrada "Senha do aparelho:" ([ref]$y)
$txtTelefone    = NovaEntrada "Telefone:" ([ref]$y)
$txtEquipamento = NovaEntrada "Equipamento:" ([ref]$y)
$txtDefeito     = NovaEntrada "Defeito apresentado:" ([ref]$y)
$txtPecas       = NovaEntrada "Peças necessárias:" ([ref]$y)

# === BOTÃO ===
$btn = New-Object System.Windows.Forms.Button
$btn.Text = "Gerar e imprimir"
$btn.Location = New-Object System.Drawing.Point(120, $y)
$btn.Width = 150
$form.Controls.Add($btn)

# === EVENTO DO BOTÃO ===
$btn.Add_Click({
    $nome = $txtNome.Text.Trim()
    $senha = $txtSenha.Text.Trim()
    $telefone = $txtTelefone.Text.Trim()
    $equipamento = $txtEquipamento.Text.Trim()
    $defeito = $txtDefeito.Text.Trim()
    $pecas = $txtPecas.Text.Trim()
    $dataHora = Get-Date -Format "dd/MM/yyyy HH:mm"

    if (-not $nome) {
        [System.Windows.Forms.MessageBox]::Show("O campo 'Nome' é obrigatório!") | Out-Null
        return
    }

    $conteudo = @"
========================================
          ORDEM DE SERVIÇO
========================================
Cliente: $nome
Telefone: $telefone
Senha: $senha
----------------------------------------
Equipamento: $equipamento
Defeito: $defeito
Peças necessárias: $pecas
----------------------------------------
Data/Hora: $dataHora
========================================
      OBRIGADO PELA PREFERÊNCIA!
========================================

"@

    # === Salvar histórico ===
    if (!(Test-Path $diretorioHistorico)) {
        New-Item -ItemType Directory -Path $diretorioHistorico | Out-Null
    }
    $arquivoNome = Join-Path $diretorioHistorico "$($nome -replace '[^\w\s-]', '') - $equipamento.txt"
    $conteudo | Out-File -FilePath $arquivoNome -Encoding UTF8 -Append

    # === Imprimir em RAW ===
    try {
        EnviarParaImpressoraRaw -printerName $nomeImpressora -text $conteudo
        [System.Windows.Forms.MessageBox]::Show("✅ Ordem salva e impressa com sucesso!") | Out-Null
    } catch {
        [System.Windows.Forms.MessageBox]::Show("❌ Erro ao imprimir: $_") | Out-Null
    }

    # Limpa os campos
    $txtNome.Clear()
    $txtSenha.Clear()
    $txtTelefone.Clear()
    $txtEquipamento.Clear()
    $txtDefeito.Clear()
    $txtPecas.Clear()
})

# === MOSTRA A JANELA ===
$form.Topmost = $true
$form.Add_Shown({$form.Activate()})
$form.ShowDialog()
