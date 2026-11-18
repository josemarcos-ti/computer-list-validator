# Validação de lista de computadores (em paralelo)

Script PowerShell que recebe uma **lista de computadores** e verifica, **em paralelo**, se cada um está **online**, o **IP**, o **domínio**, o **usuário logado**, o **sistema**, há quantos dias está **ligado** e o **espaço livre no C:** — gerando um **relatório HTML interativo** e um **CSV**.

Útil para inventário rápido, migração de domínio, auditoria de parque, conferência de lista de patrimônio ou antes de uma janela de manutenção.

![Exemplo do relatório](exemplo/screenshot.png)

> Veja o relatório de exemplo (dados fictícios): [exemplo/relatorio-exemplo.html](exemplo/relatorio-exemplo.html)

## O que ele faz

Para cada computador:
1. Resolve o nome no DNS e faz ping
2. Lê pela rede (CIM/WinRM, com tentativa via DCOM): domínio, usuário logado (inclusive sessão remota), sistema, último boot e espaço livre no C:
3. Classifica:

| Situação | Significado |
|---|---|
| **Online** | respondeu e foi lido |
| **Sem acesso** | responde na rede, mas a leitura remota falhou (permissão, firewall, WinRM) |
| **Offline** | não respondeu |
| **Sem DNS** | o nome não existe no DNS |

- Testa vários computadores ao mesmo tempo (runspaces) — listas grandes terminam em poucos minutos
- O resultado mantém a **ordem da lista original, numerada**
- Tem trava de tempo: um computador travado não segura o relatório inteiro

## Requisitos

- Windows PowerShell 5.1 ou PowerShell 7
- Acesso administrativo aos computadores (WinRM ou DCOM liberados) para ler domínio/usuário
- O modo `-Demo` não precisa de rede

## Como usar

```powershell
# lista em arquivo texto (um computador por linha)
.\Test-ComputerList.ps1 -Arquivo .\computadores.txt

# lista direto no comando, com credencial administrativa
.\Test-ComputerList.ps1 -ComputerName PC01, PC02, SRV-ARQ01 -Credential (Get-Credential)

# mais rápido para listas grandes
.\Test-ComputerList.ps1 -Arquivo .\computadores.txt -Paralelo 50

# testar sem rede (dados fictícios)
.\Test-ComputerList.ps1 -Demo
```

Exemplo de `computadores.txt`:
```
# computadores do 2º andar
FIN-PC001
FIN-PC002
RH-PC010
```

| Parâmetro | Descrição | Padrão |
|---|---|---|
| `-Arquivo` | arquivo texto, um computador por linha (`#` = comentário) | — |
| `-ComputerName` | lista de computadores no comando | — |
| `-Credential` | credencial para a leitura remota | conta atual |
| `-Paralelo` | quantos testes ao mesmo tempo | `25` |
| `-Saida` | pasta dos arquivos gerados | pasta atual |
| `-Demo` | usa dados fictícios | — |
| `-NaoAbrir` | não abre o HTML ao terminar | — |

## Observações

- O script **só lê** informações: não altera nada nos computadores.
- "Sem acesso" geralmente indica que a conta usada não é administradora na máquina ou que o WinRM/firewall está bloqueado.

## Licença

MIT — veja [LICENSE](LICENSE).
