# Organizador de Repositório

Um aplicativo com interface gráfica para organizar e gerenciar arquivos de repositório.

## Funcionalidades

- Remover arquivos PDF com padrão de nome "SEI_54202.000651_1999_01" (números variáveis)
- Remover arquivos com "ChecklistOcupante_" no nome
- Remover arquivos PDF com padrão de nome "4_PR039500000041" (números variáveis)
- Remover pastas chamadas "01_docsRecebidosEmail_Wpp"
- Mover pareceres PDF para pastas específicas de lotes

## Requisitos

- Python 3.6 ou superior
- sv-ttk (tema Sun Valley para tkinter)

## Instalação

1. Clone ou baixe este repositório
2. Instale as dependências:
   ```
   pip install -r requirements.txt
   ```

## Uso

Execute o script principal:
```
python organizador_repositorio.py
```

1. Clique em "Selecionar" para escolher o diretório que deseja organizar
2. Use os botões para executar cada ação individualmente
3. Acompanhe o progresso na barra de status

## Observações

- Cada tarefa pode ser executada independentemente
- As operações de exclusão são permanentes, então use com cuidado 