from pathlib import Path
lines = Path('qing.ps1').read_text('utf-8').splitlines()
for i in range(600, 710):
    print(f'{i+1}:{lines[i]}')
