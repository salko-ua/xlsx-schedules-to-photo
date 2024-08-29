{
    pkgs ? import <nixpkgs> { },
}:
pkgs.mkShell {
    nativeBuildInputs =
        with pkgs;
        with python312Packages;
        [
            python312
            openpyxl
            pillow
            playwright
        ];
    shellHook = ''
        export LD_LIBRARY_PATH="$LD_LIBRARY_PATH:${
            with pkgs;
            lib.makeLibraryPath [ 
                libGL 
                xorg.libX11 
                xorg.libXi 
                xorg.libXrandr
            ]
        }"
    '';
}
