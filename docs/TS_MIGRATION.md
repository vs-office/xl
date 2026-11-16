# План перевода проекта на TypeScript

Статус: фаза 1 закоммичена, фаза 2 выполнена в ветке `ts-migration` (не закоммичено), дальше фаза 3. Отклонения от плана описаны в конце файла.

Образец организации репозитория и кодовой базы: [astral-sh/ruff-vscode](https://github.com/astral-sh/ruff-vscode). Берём оттуда структуру, набор скриптов, подход к проверкам и защите цепочки поставок. Правила против `any` у нас строже (см. раздел 3).

## 0. Исходное состояние

- `src/*.ts` (хост расширения, ~1,6k строк, плоская папка) уже на TypeScript, `tsconfig` включает только `src/*.ts`.
- `out/**/*.js` (~18k строк) — исходники веб-вью на чистом JS (название папки вводит в заблуждение, это не артефакты сборки). Каждый файл это IIFE `(function (root) { const OX = root.OX || (root.OX = {}) ... })(globalThis)`, экспорт через глобальный `OX` и `module.exports`.
- `scripts/build-webview.js` склеивает файлы через terser в заданном порядке, порядок задаёт зависимости. Результат в `dist/webview/`.
- `out/vendor/` (`ssf.js`, `fflate.js`) — вендоренные библиотеки.
- Тесты `test/*.test.js` (~12k строк) подключают модули через `require('../out/core/...')`, запуск `node --test`.
- Линт: `eslint.config.js` (только JS-правила) и Prettier. CI на npm, Node 22.
- Устарело: `typescript ^4.4`, `@types/node ^11`, `webpack-cli 4`.
- Размеры: `grid.js` 3,4k, `calc.js` 3,3k, `xlsx-reader.js` 2,8k, `xlsx-writer.js` 1,9k, `app-shell.js` 1,4k, `format.js` 1,2k, `csv-provider.js` 1,2k.

## 1. Цели и не-цели

Цели:

- Весь код (хост, веб-вью, тесты) на TS со строгими настройками, без `any`.
- ES-модули вместо глобального `OX`.
- Репозиторий организован как `ruff-vscode`: `src/` как единственный корень исходников, `npm run check` как единственная команда проверки, те же имена скриптов, тот же стек инструментов.

Не-цели:

- Менять поведение, алгоритмы или производительность. Это перевод и перестановка, а не рефакторинг логики.
- Переписывать `ssf.js` (остаётся вендоренным, получает `.d.ts`).

## 2. Целевая организация репозитория (по образцу ruff-vscode)

### Структура

```
xl/
├── src/
│   ├── extension.ts            # точка входа хоста
│   ├── common/                 # общий код хоста: dispose, html, links, serializer, documentViewManager, baseDocumentView
│   ├── csv/                    # хост CSV: csvEditor, csvDocumentView, csvLineEdits
│   ├── excel/                  # хост Excel: excelEditor, excelDocumentView
│   ├── webview/                # бывший out/: код страниц (бандлится отдельно, DOM)
│   │   ├── core/               # format, csv-parser, csv-provider, xlsx-*, calc, grid-core, grid, app-shell
│   │   ├── csv.ts, excel.ts    # точки входа страниц
│   │   ├── styles/             # grid.css, vscode.css
│   │   └── vendor/             # ssf.js (+ .d.ts), LICENSE.*; fflate ставится из npm, не вендорится
│   └── test/                   # все тесты и хелперы (как src/test в ruff-vscode)
├── scripts/                    # служебные скрипты (релиз, сборка webview, если останется)
├── docs/
├── img/
├── .github/workflows/          # ci.yml, release.yml, zizmor.yml
├── .vscode/                    # launch.json, tasks.json, settings.json, extensions.json
├── .editorconfig  .gitattributes  .npmrc  .prettierrc.json  .prettierignore  .vscodeignore
├── eslint.config.mjs           # flat config, typescript-eslint + import-x
├── tsconfig.json               # один конфиг на весь src, как в ruff-vscode (lib с DOM для вебвью)
├── webpack.config.js           # массив конфигов: host, webview-excel, webview-excel-edit, webview-csv
├── package.json  package-lock.json
├── README.md  CHANGELOG.md  CONTRIBUTING.md  LICENSE
```

Точное деление `src/common` / `src/csv` / `src/excel` уточняется на фазе 1, когда будет видно, как файлы реально зависят друг от друга. Принцип: в `src/` лежат только исходники, `dist/` — только артефакты сборки, `out/` не используется (в ruff-vscode `out/` это результат `tsc` для тестов).

### Инструменты (повторяем стек ruff-vscode)

| Роль | Инструмент | Как в ruff-vscode |
|---|---|---|
| Менеджер пакетов | **npm** (остаёмся, как они) | `npm ci --ignore-scripts` в CI |
| Типы | `tsc --noEmit` | скрипт `tsc`; в devDependencies алиасы `typescript` (6.x) и `@typescript/native` (7.x) |
| Линтер | `eslint` + `typescript-eslint` + `eslint-plugin-import-x` | `eslint.config.mjs`, `--max-warnings=0` |
| Форматтер | `prettier` | `fmt-check` / `fmt` |
| Сборка | `webpack` + `ts-loader` | `webpack.config.js`, массив конфигов |
| Пакет | `@vscode/vsce` | `vsce-package` |
| Тесты | `node --test` (юнит) | у них `@vscode/test-cli` для e2e, у нас юнит-тесты остаются на `node --test` |

Решения, принятые под цель «как ruff-vscode»:

- Остаёмся на **npm** (не pnpm): так максимально похоже, и CI не переписывается.
- Остаёмся на **Prettier** (не Biome): как у них, у нас он уже есть.
- Остаёмся на **webpack + ts-loader** (не esbuild): как у них. Вебвью собирается теми же webpack-конфигами (несколько entry), вместо склейки терсером. Ленивый бандл `excel-edit` остаётся отдельным entry.

### Скрипты `package.json` (имена как в ruff-vscode)

```jsonc
{
  "scripts": {
    "check": "npm run fmt-check && npm run lint && npm run tsc && npm run test",
    "fmt": "prettier -w .",
    "fmt-check": "prettier --check .",
    "lint": "eslint src --max-warnings=0",
    "tsc": "tsc --noEmit",
    "test": "node --test src/test/*.test.ts",
    "compile": "webpack",
    "watch": "webpack --watch",
    "package": "webpack --mode production --devtool source-map --config ./webpack.config.js",
    "vsce-package": "vsce package --no-dependencies -o xl.vsix",
    "vscode:prepublish": "npm run package"
  }
}
```

### Защита цепочки поставок (как в ruff-vscode)

`.npmrc`:

```
engine-strict = true
ignore-scripts = true
min-release-age = 7
```

- `ignore-scripts` не даёт зависимостям выполнять код при установке. `min-release-age` (дни) — аналог `exclude-newer` в `uv`: не ставим версии младше 7 дней. Проверить, что установленная версия npm поддерживает этот параметр.
- CI: `permissions: {}` на уровне workflow и минимальные права на job, `persist-credentials: false` у `checkout`, действия закреплены по SHA с комментарием версии, версия Node закреплена точно, `npm ci --ignore-scripts`, `package-manager-cache: false`. Отдельный `zizmor.yml` аудирует workflow, `release.yml` собирает `.vsix` для релиза (публикация в Marketplace остаётся ручной).
- Версия Node в `engines` и в CI совпадает.

### Прочие файлы

`CONTRIBUTING.md` (как поднять окружение и запустить `npm run check`), `CHANGELOG.md`, `.gitattributes`, `.vscode/tasks.json` с npm-скриптами, `typescript.tsc.autoDetect: off` в `.vscode/settings.json`.

## 3. Правила против AnyScript

Это главное требование, и здесь мы намеренно строже, чем ruff-vscode (у них `no-explicit-any: off` и `any` встречается в коде). `any` не должен попасть в код ни одним путём.

Конфигурация (проверяет CI):

- `tsconfig`: всё из ruff-vscode (`strict`, `noImplicitReturns`, `noFallthroughCasesInSwitch`, `noUnusedParameters`) плюс `noUncheckedIndexedAccess`, `exactOptionalPropertyTypes`, `noImplicitOverride`, `isolatedModules` (`verbatimModuleSyntax` несовместим с `module: commonjs`, вместо него правило `consistent-type-imports`).
- typescript-eslint (`recommendedTypeChecked` + `strictTypeChecked`), правила как `error`: `no-explicit-any`, `no-unsafe-assignment`, `no-unsafe-member-access`, `no-unsafe-call`, `no-unsafe-return`, `no-unsafe-argument`, `no-non-null-assertion`, `ban-ts-comment` (`@ts-expect-error` только с описанием причины, `@ts-ignore` и `@ts-nocheck` запрещены), `consistent-type-assertions` (запрет `as` на литералах объектов).
- Запрет `as unknown as T` через `no-restricted-syntax`.
- Остальное из конфига ruff-vscode: `eqeqeq` (`null: never`), `no-console: error`, `import-x` recommended + typescript.
- `--max-warnings=0`: предупреждений в проекте нет.

Практические правила:

- Сначала типы данных, потом функции. Описываем `Workbook`, `Sheet`, `Cell`, `Style`, `Selection`, сообщения хост ↔ вебвью и `ZipLimits`. Всё, что имеет «вид» (`kind`/`type`), делаем дискриминированным объединением, а не объектом с опциональными полями.
- Внешние данные (содержимое xlsx/xml, `postMessage`, `JSON.parse`, настройки) приходят как `unknown` и сужаются через type guard или валидатор на границе. Дальше по коду типы честные.
- Если тип не выводится, значит модель данных описана плохо. Правим модель, а не ставим `as`.
- Запрещено «временно» ставить `any` или `// @ts-nocheck` ради зелёной сборки. Если файл не доводится, он остаётся JS (см. раздел 5).
- `Record<string, any>` и `Function` — тоже `any`. Вместо них `Record<string, unknown>` и конкретные сигнатуры.
- Метрика: `grep -rnE "\bany\b| as |@ts-" src` в итоговом дереве даёт только осмысленные исключения с комментарием.

## 4. Установка и запуск (локально, без глобальных установок)

Всё ставится как devDependencies проекта в `node_modules/`, в систему ничего. Это аналог `uv sync --dev`. Версии фиксируются в `package.json` и `package-lock.json`.

```bash
npm ci --ignore-scripts      # поднять окружение с нуля из lock-файла (как uv sync)
npm run check                # всё на всём проекте: fmt-check, lint, tsc, test
npm run fmt                  # автоформатирование
npx eslint src --fix         # автоисправление линтера
```

Однократное добавление зависимостей при настройке (фаза 1):

```bash
npm i -D typescript typescript-eslint eslint @eslint/js globals eslint-plugin-import-x \
         eslint-import-resolver-typescript prettier ts-loader webpack webpack-cli \
         @types/node @types/vscode @types/webpack-env @vscode/vsce fflate
```

`npx` и `npm run` берут бинарники из локального `node_modules/.bin`. Версию Node закрепить в `engines` и `.nvmrc`/`.node-version` (fnm у вас уже есть), это аналог `requires-python` и `.python-version`.

## 5. План действий

Каждая фаза заканчивается зелёным `npm run check` и коммитом. Остановиться можно после любой.

### Фаза 1. Каркас репозитория (без изменения логики)

1. Обновить зависимости (`typescript`, `@types/node`, `webpack-cli` и др.), заменить вендоренный `fflate` зависимостью из npm (в ruff-vscode вендоренных библиотек нет), для `ssf` написать `.d.ts`.
2. Перенести `out/` → `src/webview/`, `test/` → `src/test/`, хост-файлы разложить по `src/common`, `src/csv`, `src/excel`. Поправить пути в тестах, `.vscodeignore`, `eslint.config`.
3. `tsconfig.json` по образцу ruff-vscode, один конфиг на весь `src`, `lib` включает DOM для вебвью. Если типы DOM и API хоста начнут конфликтовать, выделяем второй конфиг (запасной вариант, не по умолчанию).
4. Webpack: массив конфигов (host + три entry вебвью) вместо `scripts/build-webview.js` с terser. Сравнить размеры бандлов до и после.
5. `eslint.config.mjs` (flat, typescript-eslint, import-x) с правилами из раздела 3. Пока на JS-файлах TS-правила точечно отключены.
6. `.npmrc`, `.gitattributes`, `CONTRIBUTING.md`, `CHANGELOG.md`, обновить `.vscode/*`.
7. CI: закрепление по SHA, `permissions: {}`, `npm ci --ignore-scripts`, Node закреплён.

Критерий выхода: сборка, `.vsix` и тесты работают как до изменений, содержимое JS ещё не тронуто.

### Фаза 2. Модули вместо `OX`

1. Для каждого файла заменить IIFE и `OX.x = ...` на `export`. Порядок по зависимостям, от листьев: `format` → `csv-parser` → `grid-core` → `xlsx-reader` → `xlsx-provider` → `csv-provider` → `calc` → `xlsx-writer` → `grid` → `app-shell` → `csv`/`excel`.
2. Точки входа и инжектируемый хостом inline-скрипт (`getOptions`, `ignoreState`) перевести на явный контракт вместо голых глобальных функций.
3. Тесты переключить на `import`.

Критерий выхода: поведение и тесты не изменились, глобального `OX` нет.

### Фаза 3. Типы, по одному модулю

Для каждого модуля в порядке из фазы 2: описать публичные типы, переименовать `.js` → `.ts`, довести до зелёного `tsc` и линтера без `any`, `npm run check`, коммит.

Порядок усилий по ценности типов: `xlsx-reader`, `xlsx-writer`, `calc`, `format`, `csv-parser`, провайдеры, `grid-core`. `grid` и `app-shell` (DOM) последними.

### Фаза 4. Тесты на TS

Тесты и хелперы (`xlsx-fixtures`, `csv-host-sim`, `csv-host-ops`) в `.ts`. Запуск `node --test` напрямую (Node 22.18+ и 24 запускают `.ts`), запасной вариант `tsx`. Механическая работа.

### Фаза 5. Закрепление

Все правила линтера включены как `error`, временные послабления убраны. Обновить `README` и `CONTRIBUTING.md`. Финальный замер: размеры бандлов, perf-тесты, время `npm run check`. Опционально: перейти на нативный `tsc` (`@typescript/native`), как у ruff-vscode, когда он стабилен.

## 6. План проверки

На каждом шаге: `npm run check` зелёный, один логический шаг на коммит.

На границах фаз дополнительно:

- **Сборка и пакет:** `npm run compile` и `npm run vsce-package` дают рабочий `.vsix`. Размеры `dist/webview/*.js` в пределах ~5% от исходных.
- **Производительность:** `grid-perf`, `xlsx-perf`, `xlsx-perf-writer`, `xlsx-bomb` укладываются в прежние бюджеты. Особое внимание ленивому бандлу `excel-edit`, ленивым shared strings и нативному inflate.
- **Ручная проверка в VS Code** (F5, Extension Development Host): `.csv`, `.tsv`, `.xlsx`, `.xlsm`, большой файл, формулы, правка и сохранение, форматы чисел, изображения. То же в web-варианте (vscode.dev или `@vscode/test-web`): бандл общий для desktop и web.
- **Round-trip xlsx:** открыть → сохранить → открыть в Excel/LibreOffice, сравнить с эталонами из тестов.
- **Чистота типов:** `grep` на `any`, `as`, `@ts-` даёт только документированные исключения.
- **Состав пакета:** содержимое `.vsix` (`vsce ls`) не содержит исходников, тестов и лишних файлов, `.vscodeignore` актуален после переезда.

## 7. Риски

| Риск | Что делаем |
|---|---|
| Переход с `OX` на модули ломает порядок инициализации и ленивую загрузку `excel-edit` | Фаза 2 отдельно и без типизации, сначала entry points, проверка вручную |
| Webpack-бандл отличается от склейки terser (минификация, `this` на верхнем уровне, размер) | Сравнить бандлы и прогнать perf-тесты до фазы 3 |
| Большие переносы файлов портят `git blame` | Переносы отдельными коммитами без правок содержимого (`git mv`) |
| `min-release-age` не поддерживается текущим npm | Проверить версию npm, при необходимости обновить Node или убрать параметр |
| Модель данных xlsx сложная, типы разрастаются | Типизировать на границе парсера, внутри узкие типы, `unknown` + guard вместо `any` |
| Рост расходов | Пилот (фаза 1 + `grid-core`) с замером `/cost`, затем пересчёт |
| `.ts` напрямую в `node --test` подходит не всегда | `erasableSyntaxOnly`, запасной вариант `tsx` |

## 8. Оценка

- Весь объём: ~$30–60 (грубо). Фаза 1 стала чуть больше из-за переносов и CI, но webpack вместо esbuild снимает риск смены сборщика.
- Объём работ: весь проект (как у ruff-vscode), но фазы 3–4 можно остановить на любом модуле.
- Пилот (фаза 1 + `grid-core`): ~$6–10, по нему уточняем оценку.

## 9. Принятые решения и открытые вопросы

Приняты (принцип: выбираем вариант, максимально похожий на ruff-vscode):

1. `fflate` берём из npm, вендоренной копии не держим. `ssf` остаётся вендоренным с `.d.ts`.
2. Один `tsconfig.json` на весь `src`. Второй конфиг только если DOM и API хоста реально конфликтуют.
3. Объём: весь проект.
4. `zizmor.yml` и `release.yml` добавляем.
5. npm, Prettier, webpack + ts-loader.

6. Хост делим на `src/common` (dispose, html, links, serializer, documentViewManager, baseDocumentView), `src/csv` и `src/excel`. Деление небольшое и нужно только ради сходства с ruff-vscode, делается `git mv` в фазе 1.
7. `CHANGELOG.md` ведём вручную, без автоматизации.
8. Публикация остаётся ручной: `.vsix` загружается в Marketplace вручную. `release.yml` только собирает `.vsix` и прикладывает его к GitHub Release (без токенов и без публикации в Marketplace/Open VSX). Публикацию из CI можно добавить позже.

Открытых вопросов нет.

## 10. Статус выполнения

### Фаза 1 (готово, не закоммичено)

- Структура: `out/` → `src/webview/`, `test/` → `src/test/`, хост разложен по `src/common`, `src/csv`, `src/excel`. `serializer.ts` оставлен в `src/` рядом с `extension.ts`, потому что он зависит и от `csv`, и от `excel`.
- Зависимости обновлены (TypeScript 6 как `typescript`, нативный 7 как `@typescript/native`, ESLint 10, typescript-eslint, import-x, webpack-cli 7). `.npmrc` с `ignore-scripts` и `min-release-age`.
- Хост (`src/**/*.ts`) проходит `strict` + `noUncheckedIndexedAccess` + `exactOptionalPropertyTypes` + `noImplicitOverride`, линт с запретом `any`, `no-unsafe-*`, `!`. `any` в хосте нет: сообщения вебвью приходят как `unknown` и сужаются через `isRecord`, добавлены типы `CsvOptions`, `ExcelOptions`, `ChunkMessage`.
- `toWorkbookBytes` вынесена в `src/excel/workbookBytes.ts`, тест импортирует её напрямую вместо вырезания кода регулярным выражением.
- CI по образцу ruff-vscode: закрепление по SHA, `permissions: {}`, `npm ci --ignore-scripts`, Node 24.20.0, `zizmor.yml`, `release.yml` (собирает `.vsix` и прикладывает к GitHub Release). Добавлены `CONTRIBUTING.md`, `CHANGELOG.md`, `.gitattributes`.
- Проверено: `npm run check` зелёный (244 теста), чистая установка из lock-файла и `vsce package` работают, состав `.vsix` тот же, кроме нового пути лицензий вендоренных библиотек.

### Отклонения от плана

- **Webpack-сборка вебвью и `fflate` из npm перенесены в фазу 2.** Пока файлы остаются IIFE с глобальными `OX`, `SSF`, `fflate`, их собирает прежний `scripts/build-webview.js` (terser). Webpack собирает только хост.
- **`verbatimModuleSyntax` не включён** (несовместим с CommonJS-хостом, у ruff-vscode его тоже нет).
- **ESLint:** `recommendedTypeChecked` плюс явные правила против `any`, а не `strictTypeChecked` (слишком шумно на коде, который переводится с JS). `eqeqeq` без `null: never` (в JS `=== null` и `== null` ведут себя по-разному, поведение не меняем).
- **Линт для старого JS мягче:** `no-useless-assignment`, `preserve-caught-error` выключены для `src/webview/**/*.js`, а `no-unused-vars`, `no-useless-escape`, `eqeqeq` для `src/test/**/*.js`. Это временно, пока файлы не станут TS.
- **`lint` запускается как `eslint .`**, а не `eslint src`, чтобы покрыть конфиги и `scripts/`.
- **Кэш npm пользователя повреждён** (root-owned файлы в `~/.npm`). Установку я делал с временным кэшем `/tmp/xl-npm-cache`. Починить: `sudo chown -R 501:20 ~/.npm`.

### Фаза 2 (готово, не закоммичено)

- Глобального `OX` больше нет. Каждый модуль в `src/webview/core` это ES-модуль с именованными экспортами (имена те же, что были у `OX.*`). `SSF` и `fflate` импортируются (`fflate` из npm, `ssf.js` остался вендоренным).
- Вебвью собирает webpack (конфиг `webview` в `webpack.config.js`), `scripts/build-webview.js` с terser удалён, остался `scripts/build-styles.js` только для CSS. Страницы `csv.js` и `excel.js` публикуют `xlPage` (`webpack library type: var`), inline-скрипт хоста вызывает `xlPage.start(options, ignoreState)` вместо глобальных `getOptions`/`ignoreState`/`handleEvents`/`initPage`.
- Ленивый чанк редактирования: `import('./edit.js')` (Calc + writer) в `excel.js`, файл `excel-edit.js`, `__webpack_public_path__` берётся из опции `scriptUri`. Формулы проверяются через опцию `validateFormula` у `XlsxProvider`, которую страница подставляет после загрузки чанка.
- `src/webview/package.json` (`type: module`) и `src/webview/vendor/package.json` (`type: commonjs`) нужны, чтобы Node загружал эти файлы без предупреждения о типе модуля.
- Тесты: заголовки переведены с `OX`-глобалей на `require` модулей, `csv-host-sim` запускает страницу в процессе теста (раньше в `vm`) с фейковым `Shell` через параметр `createShell` у `initPage`. `webview-bundle.test.js` собирает бандлы через webpack, проверяет, что `xlPage` доступен и что writer лежит только в чанке `excel-edit`.
- Проверено: `npm run check` (244 теста, 7 с), `npm run package`. Размеры production-бандлов: `excel.js` 239000 (было 239016), `excel-edit.js` 89181 (было 88270), `csv.js` 169550 (было 158579).

### Известные отклонения и риски фазы 2

- `csv.js` вырос на ~11 КБ: `format.js` теперь импортирует `ssf.js` статически, и он попал в CSV-страницу, где его раньше не было (CSV его не использует). Решение отложено на фазу 3: разделить `format.js` на ссылки на ячейки (`colName`, `cellRef`, ...) и форматирование чисел.
- В `xlsx-writer.js` добавлен тестовый шов `writerDeps.Inflate`: тест «big untouched entries are never inflated» раньше подменял `fflate.Inflate` глобально, а с модулями привязку импорта подменить нельзя.
- Не проверено вручную: загрузка чанка `excel-edit.js` в настоящем VS Code (CSP, `scriptUri`), открытие CSV/XLSX, редактирование и сохранение. Это обязательная ручная проверка перед коммитом фазы 2.
- Урок: тестовый хелпер сначала зависал, потому что страница выбирает `MessageChannel`, если он есть, а в `vm`-песочнице его не было. В хелпере теперь `delete globalThis.MessageChannel`.
