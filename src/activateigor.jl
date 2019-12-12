using Glob

igorpath = glob("Asylum*/**/Igor.exe","C:/")

targetdir = dirname(igorpath[1])
figor = joinpath(@__DIR__,"..","registerIgor6.reg")
varname = "TargetDir"
c=`powershell.exe -nop -nologo  cmd /C set $(varname)=$targetdir`
# ; cmd /C echo \%TargetDir\%`

run(c)
