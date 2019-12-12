using PyCall
using Blink
using Interact
import TableWidgets
using Tables
using CategoricalArrays
using Base.Filesystem
using TableView
win=Window()

# --
w32 = pyimport("win32com.client")
py"""
    import win32com.client as  w32
    """
igor = py"""w32.Dispatch("IgorPro.Application")"""

A=py"$(w32.constants)"
pyimport("weakref")
PyDict(igorc)
PyDict{String,Int32}(A.__dicts__[1])



for i in 1:1000
    igor.SendToHistory(0,"hey")
end

function getmethods(comobject)
    functionlist = comobject |> keys
end
w = getmethods(w32)


q=TableView.showtable(w);

body!(win,w)
w32.__file__ |> show
igor = w32.Get("IgorPro.Application.6")

igormod = w32.gencache.GetModuleForProgID("IgorPro.Application.6") |> show
