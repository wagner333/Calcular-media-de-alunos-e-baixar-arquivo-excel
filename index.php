<?php
require 'vendor/autoload.php'; // Certifique-se de que o caminho está correto

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;


$csvFile = 'dados_alunos.csv';

if ($_SERVER["REQUEST_METHOD"] == "POST" && !isset($_POST['action'])) {
    $nome = htmlspecialchars($_POST['nome']);
    $nota1_atividade = (float)htmlspecialchars($_POST['n1_atividade']);
    $nota1_prova = (float)htmlspecialchars($_POST['n1_prova']);
    $nota2_atividade = (float)htmlspecialchars($_POST['n2_atividade']);
    $nota2_prova = (float)htmlspecialchars($_POST['n2_prova']);
    $nota3_atividade = (float)htmlspecialchars($_POST['n3_atividade']);
    $nota3_prova = (float)htmlspecialchars($_POST['n3_prova']);
    $nota4_atividade = (float)htmlspecialchars($_POST['n4_atividade']);
    $nota4_prova = (float)htmlspecialchars($_POST['n4_prova']);

    // Calcula as médias dos bimestres
    $media1 = ($nota1_atividade + $nota1_prova) / 2;
    $media2 = ($nota2_atividade + $nota2_prova) / 2;
    $media3 = ($nota3_atividade + $nota3_prova) / 2;
    $media4 = ($nota4_atividade + $nota4_prova) / 2;

    // Calcula a média final
    $media_final = ($media1 + $media2 + $media3 + $media4) / 4;

    // Determina o status de reprovação
    $status = $media_final < 6 ? 'Reprovado' : 'Aprovado';

    // Adiciona os dados ao arquivo CSV
    $file = fopen($csvFile, 'a');
    fputcsv($file, [$nome, $media1, $media2, $media3, $media4, $media_final, $status]);
    fclose($file);

    $message = "Dados do aluno $nome foram salvos.";
}

if (isset($_POST['action']) && $_POST['action'] == 'generate_excel') {
    // Criação do novo arquivo Excel
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    // Adiciona cabeçalhos
    $sheet->setCellValue('A1', 'Nome');
    $sheet->setCellValue('B1', 'Média 1º Bimestre');
    $sheet->setCellValue('C1', 'Média 2º Bimestre');
    $sheet->setCellValue('D1', 'Média 3º Bimestre');
    $sheet->setCellValue('E1', 'Média 4º Bimestre');
    $sheet->setCellValue('F1', 'Média Final');
    $sheet->setCellValue('G1', 'Status');

    // Lê dados do arquivo CSV e escreve no Excel
    if (file_exists($csvFile)) {
        $row = 2;
        if (($handle = fopen($csvFile, 'r')) !== FALSE) {
            while (($data = fgetcsv($handle, 1000, ",")) !== FALSE) {
                $sheet->setCellValue('A' . $row, $data[0]);
                $sheet->setCellValue('B' . $row, $data[1]);
                $sheet->setCellValue('C' . $row, $data[2]);
                $sheet->setCellValue('D' . $row, $data[3]);
                $sheet->setCellValue('E' . $row, $data[4]);
                $sheet->setCellValue('F' . $row, $data[5]);
                $sheet->setCellValue('G' . $row, $data[6]);
                $row++;
            }
            fclose($handle);
        }
    }

    // Salva o arquivo Excel
    $writer = new Xlsx($spreadsheet);
    $filename = 'notas_alunos.xlsx';
    $writer->save($filename);

    $message = "<br><br>Os dados foram salvos no arquivo <a href=\"$filename\" download>$filename</a>.";
}
?>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Cadastro de Alunos</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-QWTKZyjpPEjISv5WaRU9OFeRpok6YctnYmDr5pNlyT2bRjXh0JMhjY6hW+ALEwIH" crossorigin="anonymous">
    <style>
        body {
            font-family: "Roboto", sans-serif;
            display: flex;
            justify-content: center;
            align-items: center;
            height: auto;
            background-color: #f8f9fa;
        }
        .container {
            background-color: #ffffff;
            padding: 2rem;
            border-radius: 0.5rem;
            box-shadow: 0 0 1rem rgba(0, 0, 0, 0.1);
        }
        .form-control {
            margin-bottom: 1rem;
        }
        .btn-submit {
            margin-top: 1rem;
        }
        .form-row {
            margin-bottom: 1rem;
        }
        .bimestre-section {
            margin-bottom: 2rem;
            padding: 1rem;
            border: 1px solid #ddd;
            border-radius: 0.5rem;
            background-color: #f9f9f9;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1 class="mb-4">Cadastro de Alunos</h1>
        <form action="" method="post" class="mb-4">
            <div class="mb-3">
                <label for="nome" class="form-label">Nome do Aluno</label>
                <input type="text" id="nome" name="nome" class="form-control" placeholder="Joãozinho" required>
            </div>

            <!-- Primeiro Bimestre -->
            <div class="bimestre-section">
                <h4 class="mb-3">Primeiro Bimestre</h4>
                <div class="form-row row">
                    <div class="col-md-6 mb-3">
                        <label for="n1_atividade" class="form-label">Nota Atividade</label>
                        <input type="text" id="n1_atividade" name="n1_atividade" class="form-control" placeholder="10" maxlength="3" required>
                    </div>
                    <div class="col-md-6 mb-3">
                        <label for="n1_prova" class="form-label">Nota Prova</label>
                        <input type="text" id="n1_prova" name="n1_prova" class="form-control" placeholder="10" maxlength="3" required>
                    </div>
                </div>
            </div>

            <!-- Segundo Bimestre -->
            <div class="bimestre-section">
                <h4 class="mb-3">Segundo Bimestre</h4>
                <div class="form-row row">
                    <div class="col-md-6 mb-3">
                        <label for="n2_atividade" class="form-label">Nota Atividade</label>
                        <input type="text" id="n2_atividade" name="n2_atividade" class="form-control" placeholder="10" maxlength="3" required>
                    </div>
                    <div class="col-md-6 mb-3">
                        <label for="n2_prova" class="form-label">Nota Prova</label>
                        <input type="text" id="n2_prova" name="n2_prova" class="form-control" placeholder="10" maxlength="3" required>
                    </div>
                </div>
            </div>

            <!-- Terceiro Bimestre -->
            <div class="bimestre-section">
                <h4 class="mb-3">Terceiro Bimestre</h4>
                <div class="form-row row">
                    <div class="col-md-6 mb-3">
                        <label for="n3_atividade" class="form-label">Nota Atividade</label>
                        <input type="text" id="n3_atividade" name="n3_atividade" class="form-control" placeholder="10" maxlength="3" required>
                    </div>
                    <div class="col-md-6 mb-3">
                        <label for="n3_prova" class="form-label">Nota Prova</label>
                        <input type="text" id="n3_prova" name="n3_prova" class="form-control" placeholder="10" maxlength="3" required>
                    </div>
                </div>
            </div>

            <!-- Quarto Bimestre -->
            <div class="bimestre-section">
                <h4 class="mb-3">Quarto Bimestre</h4>
                <div class="form-row row">
                    <div class="col-md-6 mb-3">
                        <label for="n4_atividade" class="form-label">Nota Atividade</label>
                        <input type="text" id="n4_atividade" name="n4_atividade" class="form-control" placeholder="10" maxlength="3" required>
                    </div>
                    <div class="col-md-6 mb-3">
                        <label for="n4_prova" class="form-label">Nota Prova</label>
                        <input type="text" id="n4_prova" name="n4_prova" class="form-control" placeholder="10" maxlength="3" required>
                    </div>
                </div>
            </div>

            <input type="submit" value="Cadastrar Aluno" class="btn btn-primary btn-submit">
        </form>

        <form action="" method="post">
            <input type="hidden" name="action" value="generate_excel">
            <input type="submit" value="Gerar Excel com Todos os Dados" class="btn btn-success">
        </form>

        <?php if (isset($message)) echo $message; ?>
    </div>
    <script src="https://kit.fontawesome.com/2fd24895a9.js" crossorigin="anonymous"></script>   
    <script src="https://cdn.jsdelivr.net/npm/@popperjs/core@2.11.8/dist/umd/popper.min.js" integrity="sha384-I7E8VVD/ismYTF4hNIPjVp/Zjvgyol6VFvRkX/vR+Vc4jQkC+hVqc2pM8ODewa9r" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.min.js" integrity="sha384-0pUGZvbkm6XF6gxjEnlmuGrJXVbNuzT9qBBavbLwCsOGabYfZo0T0to5eqruptLy" crossorigin="anonymous"></script>
</body>
</html>

