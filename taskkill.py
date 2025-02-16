import subprocess
import os

class ProcessKiller:
    def __init__(self, processos: str):
        self.processos = processos.split(',')
        self.usuario_atual = os.getlogin()

    def kill_processes(self):
        """
            Terminates specific processes associated with the current user.

            This method iterates through a list of process names, standardizes their format 
            (by appending `.exe` if necessary), and attempts to terminate them using the 
            `taskkill` command. The process termination is filtered to apply only to the 
            current user.

            Processing:
                - Strips and standardizes each process name in the `self.processos` list.
                - Appends `.exe` to the process name if it does not already end with `.exe`.
                - Uses the `taskkill` command with specific flags to terminate the processes 
                for the current user.
                - Silences the output from the `subprocess.run` call to keep execution clean.
        """
        for nome_processo in self.processos:
            nome_processo = nome_processo.strip()
            if not nome_processo.lower().endswith('.exe'):
                nome_processo += '.exe'
            try:
                print(f"Processo: {nome_processo} - Usuário: {self.usuario_atual}")
                subprocess.run(
                    [
                        "taskkill",
                        "/F",  # Força o encerramento
                        "/T",  # Encerra os processos filhos
                        "/IM", f"{nome_processo}*",  # Nome do processo (com wildcard *)
                        "/FI", f"USERNAME eq {self.usuario_atual}"  # Filtro para o usuário atual
                    ],
                    check=True,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL
                )
            except subprocess.CalledProcessError:
                pass
        return "Finalizado"

def main(processos):
    killer = ProcessKiller(processos)
    return killer.kill_processes()	

# if __name__ == "__main__":
#     inProcessos = "msedge,notepad,excel"
#     main(inProcessos)
