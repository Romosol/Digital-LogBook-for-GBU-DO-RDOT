import http from 'node:http';

const PORT = process.env.PORT || 3000;

const htmlContent = `<!DOCTYPE html>
<html lang="ru">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Digital LogBook for GBU DO RDOT</title>
  <meta name="description" content="Электронный журнал учёта работы педагога ДОП">
  <meta property="og:title" content="Digital LogBook for GBU DO RDOT">
  <meta property="og:description" content="Электронный журнал учёта работы педагога ДОП">
  <style>
    :root {
      --bg: #f8fafc;
      --card-bg: #ffffff;
      --text: #0f172a;
      --text-muted: #64748b;
      --primary: #0284c7;
      --primary-hover: #0369a1;
      --border: #e2e8f0;
      --success: #10b981;
    }
    body {
      margin: 0;
      padding: 0;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
      background: var(--bg);
      color: var(--text);
      display: flex;
      justify-content: center;
      align-items: center;
      min-height: 100vh;
      box-sizing: border-box;
    }
    .container {
      max-width: 680px;
      margin: 24px;
      padding: 32px;
      background: var(--card-bg);
      border-radius: 12px;
      border: 1px solid var(--border);
      box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -2px rgba(0, 0, 0, 0.05);
    }
    .badge {
      display: inline-block;
      padding: 4px 10px;
      border-radius: 9999px;
      background: #ecfdf5;
      color: #047857;
      font-size: 12px;
      font-weight: 600;
      margin-bottom: 12px;
    }
    h1 {
      margin: 0 0 8px 0;
      font-size: 24px;
      color: #0f172a;
    }
    p {
      color: var(--text-muted);
      line-height: 1.6;
      margin: 0 0 20px 0;
    }
    .card {
      background: #f1f5f9;
      border-radius: 8px;
      padding: 16px;
      margin-bottom: 20px;
      border: 1px solid #e2e8f0;
    }
    .card h3 {
      margin: 0 0 8px 0;
      font-size: 15px;
      color: #1e293b;
    }
    .card ul {
      margin: 0;
      padding-left: 20px;
      color: #334155;
      font-size: 14px;
      line-height: 1.7;
    }
    .code-block {
      background: #1e293b;
      color: #f8fafc;
      padding: 12px 16px;
      border-radius: 6px;
      font-family: monospace;
      font-size: 13px;
      margin: 12px 0;
      overflow-x: auto;
    }
    .btn {
      display: inline-block;
      background: var(--primary);
      color: #ffffff;
      padding: 10px 18px;
      border-radius: 6px;
      text-decoration: none;
      font-size: 14px;
      font-weight: 600;
      transition: background 0.15s ease;
    }
    .btn:hover {
      background: var(--primary-hover);
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="badge">● Сервер активен — Версия 1.0.5</div>
    <h1>Журнал учёта работы педагога ДОП</h1>
    <p>
      Настольное графическое приложение на Python (Tkinter + openpyxl) для автоматизированного
      заполнения и формирования печатного бланка классного журнала (40 страниц журнала в книге Excel).
    </p>

    <div class="card">
      <h3>Запуск на компьютере:</h3>
      <ul>
        <li>Скачайте и запустите автономный файл <strong>Journal_DOP.exe</strong> из релизов GitHub.</li>
        <li>Или запустите через Python:
          <div class="code-block">python app_gui.py</div>
        </li>
      </ul>
    </div>

    <div class="card">
      <h3>Структура 40 страниц журнала:</h3>
      <ul>
        <li><strong>Стр. 1–3:</strong> Титульный лист, Оборот (правила), Основные данные</li>
        <li><strong>Стр. 4–27:</strong> 12 учебных месяцев (Сентябрь — Август)</li>
        <li><strong>Стр. 28–29:</strong> Учёт массовых мероприятий</li>
        <li><strong>Стр. 30–31:</strong> Творческие достижения обучающихся</li>
        <li><strong>Стр. 32–37:</strong> Список обучающихся (3 разворота по 10 чел.)</li>
        <li><strong>Стр. 38–39:</strong> Инструктаж по технике безопасности</li>
        <li><strong>Стр. 40:</strong> Годовой цифровой отчёт</li>
      </ul>
    </div>

    <a href="https://github.com/Romosol/Digital-LogBook-for-GBU-DO-RDOT/releases" target="_blank" rel="noopener noreferrer" class="btn">
      Страница релизов на GitHub ➔
    </a>
  </div>
</body>
</html>`;

const server = http.createServer((req, res) => {
  res.writeHead(200, {
    'Content-Type': 'text/html; charset=utf-8',
    'Cache-Control': 'no-cache'
  });
  res.end(htmlContent);
});

server.listen(PORT, '0.0.0.0', () => {
  console.log(`Logbook dev server is running on http://0.0.0.0:${PORT}`);
});
