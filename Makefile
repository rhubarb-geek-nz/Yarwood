# Copyright (c) 2024 Roger Brown.
# Licensed under the MIT License.

BINDIR=bin\$(VSCMD_ARG_TGT_ARCH)
APPNAME=dispapp
DLLNAME=displib
TLBNAME=disptlb
LIBNAME=RhubarbGeekNzYarwood

all: $(BINDIR) $(BINDIR)\$(LIBNAME).dll $(TLBNAME)\$(BINDIR)\$(LIBNAME).dll $(BINDIR)\$(APPNAME).exe

clean:
	if exist $(BINDIR) rmdir /q /s $(BINDIR)
	cd $(DLLNAME)
	$(MAKE) clean
	cd ..
	cd $(APPNAME)
	$(MAKE) clean
	cd ..
	cd $(TLBNAME)
	$(MAKE) clean
	cd ..

$(BINDIR):
	mkdir $@

$(BINDIR)\$(APPNAME).exe: $(APPNAME)\$(BINDIR)\$(APPNAME).exe
	copy $(APPNAME)\$(BINDIR)\$(APPNAME).exe $@

$(BINDIR)\$(LIBNAME).dll: $(DLLNAME)\$(BINDIR)\$(DLLNAME).dll
	copy $(DLLNAME)\$(BINDIR)\$(DLLNAME).dll $@

$(APPNAME)\$(BINDIR)\$(APPNAME).exe:
	cd $(APPNAME)
	$(MAKE) CertificateThumbprint=$(CertificateThumbprint)
	cd ..

$(DLLNAME)\$(BINDIR)\$(DLLNAME).dll:
	cd $(DLLNAME)
	$(MAKE) CertificateThumbprint=$(CertificateThumbprint)
	cd ..

$(TLBNAME)\$(BINDIR)\$(LIBNAME).dll:
	cd $(TLBNAME)
	$(MAKE) CertificateThumbprint=$(CertificateThumbprint)
	cd ..
