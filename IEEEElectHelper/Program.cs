// See https://aka.ms/new-console-template for more information

using IEEEElectHelper;

//Loader.Test();
Loader.Testagain();
var contents = File.ReadAllLines("assets\\results\\skipped\\skipped.txt");
Array.Sort(contents);
var con = Functions.RemoveEmptyLines(contents);
File.WriteAllLines("assets\\results\\skipped\\skipped.txt", con);
//Loader.check();
