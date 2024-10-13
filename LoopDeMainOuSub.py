  i = 1
  bError = False
  while i <= 3:
      try:
          print("Tentativa", i)
          bError = False
      except Exception as e:
          print(f"Ocorreu um erro: {e}")
          if i == 3:
              print("Precisa enviar email do erro")
          bError = True
      finally:
          print("Execução da tentativa", i, "finalizada")
          if bError and i != 3:
              print("Matando Processos")
              i += 1
              continue
          print("Matando Processos")
          break
