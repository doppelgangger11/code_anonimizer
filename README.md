# code_anonimizer

Тестовая версия 

# 1. Сначала сухой прогон — посмотреть, какие колонки скрипт считает "чувствительными"

python project_anonymizer.py analyze /path/to/project

# 2. Анонимизация (создаёт новую папку, оригинал не трогает)

python project_anonymizer.py anonymize /path/to/project
    --output /path/to/project_anonymized
    --map-file /path/to/anonymization_map.json
    --columns "Компания,Клиент,Менеджер"
    --anonymize-filenames
    --overwrite
