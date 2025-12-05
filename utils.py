# utils.py
from tkinter import messagebox

def alerta(titulo, mensagem):
    messagebox.showinfo(titulo, mensagem)

def erro(mensagem):
    messagebox.showerror("Erro", mensagem)

def confirmar(mensagem):
    return messagebox.askyesno("Confirmar", mensagem)
