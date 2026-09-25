# Building DocWeft on Windows

Пошаговый гайд: полная сборка + тесты + пример (`basic_report.exe` + `template.docx` + `report.docx`), идентичная тому, что уже проверено на Linux.

Поддерживаются **обе разрядности**: `x64` и `x86`. Команды отличаются только одним флагом.

---

## 1. Предварительные требования

Установите один раз:

| Компонент | Версия | Как установить |
|---|---|---|
| **Visual Studio 2022** | любой SKU, включая Community | [visualstudio.microsoft.com](https://visualstudio.microsoft.com/) — рабочая нагрузка **"Desktop development with C++"**. Обязательно включите компонент *"C++ CMake tools for Windows"*. |
| **Git** | 2.40+ | [git-scm.com](https://git-scm.com/download/win) |
| **vcpkg** | любой актуальный | см. ниже |

### Установка vcpkg (один раз на машину)

Откройте **PowerShell** (не "x64 Native Tools", обычный PowerShell):

```powershell
cd C:\
git clone https://github.com/microsoft/vcpkg.git
.\vcpkg\bootstrap-vcpkg.bat

# Сохраняем путь в переменной среды (для всех будущих сессий)
[Environment]::SetEnvironmentVariable('VCPKG_ROOT', 'C:\vcpkg', 'User')
```

Перезапустите PowerShell, чтобы `$env:VCPKG_ROOT` стал виден.

Проверка:
```powershell
echo $env:VCPKG_ROOT
# → C:\vcpkg
```

---

## 2. Клонирование и конфигурация

```powershell
cd C:\src
git clone <url-репозитория> DocWeft
cd DocWeft
```

### Вариант A — через CMakePresets (рекомендуется)

```powershell
# x64
cmake --preset win-x64-vcpkg

# x86
cmake --preset win-x86-vcpkg
```

Первый запуск качает зависимости из vcpkg (`pugixml`, `libzip`, `nlohmann-json`, `fmt`, `inja`, `catch2`) — **15–30 минут** на обычном SSD + домашнем интернете. Последующие запуски используют binary cache vcpkg и конфигурятся за секунды.

### Вариант B — без presets (CMake < 3.25 или ручной контроль)

```powershell
# x64
cmake -B build\win-x64 `
  -G "Visual Studio 17 2022" -A x64 `
  -DCMAKE_TOOLCHAIN_FILE="$env:VCPKG_ROOT\scripts\buildsystems\vcpkg.cmake" `
  -DVCPKG_TARGET_TRIPLET=x64-windows `
  -DDOCWEFT_BUILD_SHARED=ON `
  -DDOCWEFT_BUILD_TESTS=ON `
  -DDOCWEFT_BUILD_EXAMPLES=ON `
  -DDOCWEFT_USE_FETCHCONTENT=OFF

# x86
cmake -B build\win-x86 `
  -G "Visual Studio 17 2022" -A Win32 `
  -DCMAKE_TOOLCHAIN_FILE="$env:VCPKG_ROOT\scripts\buildsystems\vcpkg.cmake" `
  -DVCPKG_TARGET_TRIPLET=x86-windows `
  -DDOCWEFT_BUILD_SHARED=ON `
  -DDOCWEFT_BUILD_TESTS=ON `
  -DDOCWEFT_BUILD_EXAMPLES=ON `
  -DDOCWEFT_USE_FETCHCONTENT=OFF
```

---

## 3. Сборка

### Через presets

```powershell
cmake --build --preset win-x64-vcpkg    # или win-x86-vcpkg
```

### Без presets

```powershell
cmake --build build\win-x64 --config Release -j
cmake --build build\win-x86 --config Release -j
```

Артефакты:

```
build\win-x64\Release\docweft.dll         # главная библиотека (SHARED)
build\win-x64\Release\docweft.lib         # import-lib для линковки
build\win-x64\tests\Release\docweft_tests.exe
build\win-x64\examples\Release\basic_report.exe
build\win-x64\examples\Release\generate_template.exe
build\win-x64\examples\Release\template.docx  ← сгенерирован автоматически
build\win-x64\examples\Release\logo.png       ← скопирован из examples\logo.png (75 B, 2×2 PNG)
build\win-x64\examples\Release\docweft.dll ← скопирован рядом с basic_report.exe
build\win-x64\examples\Release\pugixml.dll    ← транзитивные DLL (если vcpkg-порт shared)
build\win-x64\examples\Release\zip.dll
```

CMake автоматически копирует `docweft.dll` и все транзитивные runtime-DLL рядом с `basic_report.exe` (см. `examples/CMakeLists.txt`). Ничего руками делать не нужно.

---

## 4. Запуск тестов

```powershell
# Через preset
ctest --preset win-x64-vcpkg

# Без preset
ctest --test-dir build\win-x64 -C Release --output-on-failure
```

Ожидаемый результат (ровно как на Linux):
```
100% tests passed, 0 tests failed out of 28

Total Test time (real) =   0.15 sec
```

> **Почему тесты не ломаются без PATH'а?**
> `docweft_tests.exe` линкуется со *статическим* зеркалом `docweft_internal.lib` (см. `CMakeLists.txt` §"Internal static twin"). Это сделано специально: тесты нуждаются во внутренних символах (`DocxMerger`, `docx::Reader`), которые SHARED-`docweft.dll` не экспортирует (`visibility=hidden` + `DOCWEFT_API`).

---

## 5. Запуск примера

```powershell
cd build\win-x64\examples\Release
.\basic_report.exe
```

Вывод (если LibreOffice установлен и виден через `PATH` или `C:\Program Files\LibreOffice\program\soffice.exe`):
```
DocWeft v0.1.0
  template: template.docx
  output:   report.docx

Wrote report.docx
Open it in MS Word / LibreOffice to verify:
  - title 'Q1 2026 Summary'
  - author 'Иван Петров'
  - today's date (2026-04-22)
  - greeting 'Hello Alice, welcome to Acme Corp.'
  - Scores table with 3 rows:
      Alice — 95
      Борис — 82
      Carol Жу — 77
  - embedded PNG logo from logo.png (via setImage, upscaled to 64×64 via ImageSize{min=64, max=256})
Wrote report.pdf (PDF via Word or LibreOffice, whichever was found)
Wrote report.html (HTML via Word or LibreOffice, whichever was found)
```

Если LibreOffice не установлен — `.pdf`/`.html` пропускаются с подсказкой, но `report.docx` всё равно записывается и валиден.

Откройте `report.docx` в Word — все 4 заголовочные подстановки должны быть на месте, таблица содержит 1 заголовочную + 3 data-строки (шаблонная строка с `{{rowName}}`/`{{rowScore}}` удалена), форматирование (жирный, границы ячеек) сохранено, кириллица корректно отображается. `report.pdf` должен визуально совпадать с `.docx`-версией.

---

## 6. Полная однокоманда: x64 + x86 + тесты + пример

PowerShell-скрипт:

```powershell
foreach ($arch in 'x64', 'x86') {
    $preset = "win-$arch-vcpkg"
    cmake --preset $preset
    cmake --build --preset $preset
    ctest --preset $preset
    & "build\$preset\examples\Release\basic_report.exe"
}
```

---

## 7. Частые проблемы

### `CMake Error: Could not find toolchain file`

```
CMAKE_TOOLCHAIN_FILE: C:\vcpkg\scripts\buildsystems\vcpkg.cmake
```

Проверьте, что `$env:VCPKG_ROOT` установлен **в пользовательских переменных среды** (System Properties → Environment Variables), а не только в текущей сессии. Перезапустите PowerShell.

### `error MSB8020: The build tools for v143 cannot be found`

Не установлена нагрузка *"Desktop development with C++"*. Запустите Visual Studio Installer → Modify → включите компонент.

### Первая сборка зависает на `Installing ...`

Это vcpkg собирает зависимости из исходников. Прогресс виден в `build\<preset>\vcpkg-manifest-install.log`. В среднем 15–30 минут первый раз, затем вcpkg кэширует бинарники в `%LOCALAPPDATA%\vcpkg\archives\`.

### `The code execution cannot proceed because docweft.dll was not found`

Вы запускаете `basic_report.exe` из другой директории. Либо:
1. `cd` в папку с exe-файлом (CMake копирует DLL туда автоматически).
2. Или добавьте путь к `docweft.dll` в PATH:
   ```powershell
   $env:PATH = "build\win-x64\Release;$env:PATH"
   .\build\win-x64\examples\Release\basic_report.exe
   ```

### Ошибки с кириллицей в consoleoutput

PowerShell 7+ использует UTF-8 по умолчанию. Для `cmd.exe`/старых PowerShell:
```powershell
chcp 65001
```

---

## 8. Параллель между Linux и Windows

| Шаг | Linux | Windows |
|---|---|---|
| Configure | `cmake --preset linux-x64` | `cmake --preset win-x64-vcpkg` |
| Build | `cmake --build --preset linux-x64` | `cmake --build --preset win-x64-vcpkg` |
| Test | `ctest --preset linux-x64` | `ctest --preset win-x64-vcpkg` |
| Run example | `cd build/.../examples && LD_LIBRARY_PATH=.. ./basic_report` | `cd build\...\examples\Release && .\basic_report.exe` |
| Артефакт либы | `libdocweft.so.0.1.0` | `docweft.dll` + `docweft.lib` |
| Package manager | Nix / FetchContent | vcpkg |
| Lookup зависимостей | `rpath` / `LD_LIBRARY_PATH` | PATH / DLL рядом с exe |

Суть пайплайна одинаковая, отличается только упаковка артефактов.
