# AUTO ATESTADO
Parceria com [Samuel Batista](https://github.com/Samuel-Batista)

## Como usar:

### 1. Arquivos necessários:
- `AutoAtestado.exe` (ou pasta `AutoAtestado` com o executável)
- `atestados.xlsx` (planilha com os dados dos alunos)
- `chromedriver.exe` (driver do Chrome para automação)

### 2. Preparação da planilha:
Abra o arquivo `atestados.xlsx` e preencha os dados na planilha "Plan1":
- **Coluna A**: Nome do aluno (opcional, apenas para referência)
- **Coluna B**: ID do aluno (obrigatório)
- **Coluna C**: Data de início do atestado (formato: DD/MM/AAAA)
- **Coluna D**: Data de fim do atestado (formato: DD/MM/AAAA)

### 3. Execução:
2. **Opção 1**: Execute diretamente o `AutoAtestado.exe`

### 4. Durante a execução:
- O programa solicitará seu usuário e senha do sistema SENAC
- Aguarde o processamento de cada aluno
- Um arquivo de log será criado na pasta `log` com o registro de todas as operações

### 5. Logs:
- Os logs são salvos automaticamente na pasta `log`
- Cada execução gera um novo arquivo com data e hora
- Os logs contêm informações detalhadas sobre cada processamento

### 6. Observações importantes:
- Mantenha o Chrome fechado durante a execução
- Não mova o mouse ou teclado durante o processamento
- Certifique-se de ter uma conexão estável com a internet
- O programa funciona apenas com cursos "EMÉDIO 2025"

### 7. Estrutura de arquivos:
```
pasta_do_programa/
├── AutoAtestado.exe (ou pasta AutoAtestado/)
├── atestados.xlsx
├── chromedriver.exe
├── README.md
└── log/ (criada automaticamente)
    └── log_DD_MM_AA__HH_MMh.txt
```

### 8. Solução de problemas:
- **Erro "chromedriver não encontrado"**: Certifique-se de que o chromedriver.exe está na mesma pasta
- **Erro "planilha não encontrada"**: Verifique se o arquivo atestados.xlsx está na pasta correta
- **Erro de login**: Verifique suas credenciais do sistema SENAC
- **Chrome não abre**: Certifique-se de que o Chrome está instalado e atualizado

### 9. Suporte:
Em caso de problemas, verifique o arquivo de log na pasta `log` para mais detalhes sobre o erro.
