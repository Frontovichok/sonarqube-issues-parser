## Инструмент для парсинга ошибок из SonarQube
Загружает из SonarQube API информацию об ошибках и формирует отчеты в форматах xlsx и html
Поиск выполняется для ошибок из категорий Security и Reliability. Maintainability игнорируется т.к. не содержит уязвимости и бага, а в основном "плохо пахнущий код"

Конфигурация проекта, который будет парситься задается в конфигурационном файле формата JSON:
```JSON
{
    "url": "https://sonarqube.dev",
	"project_id": "id",
	"project_name": "name",
	"project_version": "verion",
	"branch": "branch",
    "JWT-SESSION": "JWT-session"
}
```


## Запуск:
```bash
python parser.py <config.json>
```
## Вывод:
Файл в текущей директории
- sonarqube_issues_report.xlsx
- sonarqube_comprehensive_report.html
- response_output.json

##
Для запуска необходимо наличие дополнительных библиотек:
- requests
- openpyxl
- urllib3
- beautifulsoup4

которые можно установить командой:
```bash
pip install -r requirements.txt
```


