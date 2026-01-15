# Toolbox
## Instalação

Dentro do diretório onde há o pyproject.toml

Para instalação
```python -m pip install .```

Para instalação em ambiente de desenvolvimento 
```python -m pip install -e .```

## Execução

Dentro do diretório onde há um arquivo de input.xlxs

Para ver os possíveis comandos
```python -m toolbox -h```
ou
```python -m toolbox --help```

Atualmente existem três modos de execução do Toolbox
1. Atualização de Bases
```python -m toolbox -a``` 
ou
```python -m toolbox --atualizar```

2. Execução de comandos via input.xlsx
```python -m toolbox -e```
ou
```python -m toolbox --executar```