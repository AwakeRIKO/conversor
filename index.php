<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Conversor PDF para Excel</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            background-color: #121212;
            color: #fff;
            margin: 0;
            padding: 0;
            display: flex;
            justify-content: center;
            align-items: center;
            height: 100vh;
        }

        .container {
            text-align: center;
            max-width: 600px;
        }

        h1 {
            font-size: 2rem;
            margin-bottom: 20px;
        }

        form {
            background-color: #1e1e1e;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.3);
        }

        input[type="file"] {
            margin: 10px 0;
            padding: 10px;
            border: none;
            border-radius: 5px;
            background-color: #2e2e2e;
            color: #fff;
        }

        input[type="submit"] {
            background-color: #ffc107;
            color: #000;
            padding: 10px 20px;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            font-weight: bold;
        }

        input[type="submit"]:hover {
            background-color: #e0a800;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>Converter PDF para Excel</h1>
        <form action="upload.php" method="post" enctype="multipart/form-data">
            <p>Selecione um arquivo PDF para converter:</p>
            <input type="file" name="file" accept=".pdf" required>
            <br>
            <input type="submit" value="Converter">
        </form>
    </div>
</body>
</html>