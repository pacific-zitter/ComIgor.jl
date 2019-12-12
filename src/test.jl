using PyCall
using Blink


igpath = joinpath("C:/AsylumResearch","Igor Pro Folder","Igor.exe")
w32 = pyimport("win32com.client")
w32.gencache.EnsureDispatch("IgorPro.Application");

igor = w32.Dispatch("IgorPro.Application")
io = IOBuffer()
igor.Status1
igor.Visible
q= propertynames(igor) .|> x->println(x)
y |> typeof

igor.Visible = true

# PyCall.pyptr_query(igo)
PyCall.python_cmd(`makepy`) |> run


y = w32.selecttlb.SelectTlb()

y.Resolve()

igortlb=getproperty(y,:dll)

r = w32.pythoncom

ig=r.LoadTypeLib(igortlb)

q=r |> propertynames .|> String

qq = q .* '\n'

HTML(join(qq))
w = Window()

Blink.AtomShell.install()
