while (true) {
        try {
            new ActiveXObject("Scripting.FileSystemObject").GetStandardStream(1).Write(">>> ");
            new ActiveXObject("Scripting.FileSystemObject").GetStandardStream(1).WriteLine(eval(new ActiveXObject("Scripting.FileSystemObject").GetStandardStream(0).ReadLine()));
        } catch (e) {
            new ActiveXObject("Scripting.FileSystemObject").GetStandardStream(2).WriteLine((e.number + 0x100000000).toString(16) + " " + e.name + ": " + e.description);
        }
}
