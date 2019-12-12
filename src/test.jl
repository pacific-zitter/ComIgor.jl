using PyCall
using Blink
using Interact
import TableWidgets
using Tables
using CategoricalArrays
using Base.Filesystem
using TableView
win=Window()
Interact.
# --
w32 = pyimport("win32com.client")
ics = pyimport("win32com.client.constants")
w32.
CategoricalArray(w,ordered=true)

function getmethods(comobject)
    functionlist = comobject |> propertynames |>  x -> OrderedDict(gensym.(x).=>x)
end

w = getmethods(w32)


q=TableView.showtable(w);

body!(win,w)
w32.__file__ |> show
igor = w32.Dispatch("IgorPro.Application.6")

igormod = w32.gencache.GetModuleForProgID("IgorPro.Application.6") |> show



r=getmethods(igormod)
showtable(r)

q=igormod."Application"

r=q()
r.__dir__()
