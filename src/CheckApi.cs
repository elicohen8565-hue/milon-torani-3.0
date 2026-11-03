// Developer tool (not part of the product).
// Checks a FINISHED program: every method, field and type it uses from the .NET libraries must exist in the
// .NET Framework 4.5 reference assemblies - that is what Windows really has. A program built against Mono's own
// libraries can call methods that exist only there (3.0.0.3: String.Split(char, StringSplitOptions)); it runs under Mono,
// passes the self-test, and dies on Windows with MissingMethodException.
//   mono CheckApi.exe TorahDictionary.exe              prints "API CHECK PASSED" (exit 0) or the missing members (exit 1)
// Build line: see build.sh (needs Mono.Cecil, which ships with Mono).
using System;
using System.Collections.Generic;
using Mono.Cecil;
using Mono.Cecil.Cil;

static class CheckApi
{
    const string RefDir = "/usr/lib/mono/4.5-api";

    static IEnumerable<TypeDefinition> AllTypes(IEnumerable<TypeDefinition> types)
    {
        foreach (TypeDefinition t in types)
        {
            yield return t;
            foreach (TypeDefinition n in AllTypes(t.NestedTypes)) yield return n;
        }
    }

    static bool External(ModuleDefinition m, TypeReference t)
    {
        while (t is TypeSpecification) t = ((TypeSpecification)t).ElementType;   // arrays, generics, by-ref...
        if (t is GenericParameter) return false;
        return t.Scope != null && t.Scope != m && !(t.Scope is ModuleDefinition);
    }

    static int Main(string[] args)
    {
        if (args.Length < 1) { Console.WriteLine("usage: CheckApi.exe <program.exe>"); return 2; }
        DefaultAssemblyResolver resolver = new DefaultAssemblyResolver();
        resolver.AddSearchDirectory(RefDir);
        ReaderParameters rp = new ReaderParameters();
        rp.AssemblyResolver = resolver;
        ModuleDefinition m = ModuleDefinition.ReadModule(args[0], rp);

        int checkedCount = 0;
        SortedDictionary<string, string> missing = new SortedDictionary<string, string>();
        HashSet<string> seen = new HashSet<string>();

        Action<MemberReference, string> check = delegate (MemberReference mr, string where)
        {
            TypeReference decl = mr is TypeReference ? (TypeReference)mr : mr.DeclaringType;
            if (decl == null || !External(m, decl)) return;
            string key = mr.FullName;
            if (!seen.Add(key)) return;
            checkedCount++;
            bool ok;
            try
            {
                if (mr is MethodReference) ok = ((MethodReference)mr).Resolve() != null;
                else if (mr is FieldReference) ok = ((FieldReference)mr).Resolve() != null;
                else { TypeReference tr = mr as TypeReference; ok = tr.Resolve() != null; }
            }
            catch (Exception) { ok = false; }
            if (!ok) missing[key] = where;
        };

        foreach (CustomAttribute ca in m.Assembly.CustomAttributes) check(ca.Constructor, "assembly attribute");

        foreach (TypeDefinition t in AllTypes(m.Types))
        {
            if (t.BaseType != null) check(t.BaseType, t.FullName + " (base type)");
            foreach (InterfaceImplementation i in t.Interfaces) check(i.InterfaceType, t.FullName + " (interface)");
            foreach (FieldDefinition f in t.Fields) check(f.FieldType, t.FullName + "::" + f.Name + " (field type)");
            foreach (MethodDefinition md in t.Methods)
            {
                check(md.ReturnType, t.FullName + "::" + md.Name + " (return type)");
                foreach (ParameterDefinition p in md.Parameters) check(p.ParameterType, t.FullName + "::" + md.Name + " (parameter)");
                if (!md.HasBody) continue;
                foreach (VariableDefinition v in md.Body.Variables) check(v.VariableType, t.FullName + "::" + md.Name + " (local)");
                foreach (Instruction ins in md.Body.Instructions)
                {
                    MemberReference mr = ins.Operand as MemberReference;
                    if (mr != null) check(mr, t.FullName + "::" + md.Name);
                }
            }
        }

        Console.WriteLine("checked " + checkedCount + " different members/types of the .NET libraries against " + RefDir);
        if (missing.Count == 0) { Console.WriteLine("API CHECK PASSED"); return 0; }
        foreach (KeyValuePair<string, string> kv in missing) Console.WriteLine("  MISSING in .NET Framework 4.5: " + kv.Key + "   (used in " + kv.Value + ")");
        Console.WriteLine("API CHECK FAILED (" + missing.Count + ")");
        return 1;
    }
}
