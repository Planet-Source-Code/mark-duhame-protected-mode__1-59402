Title:	 Protected Mode
Purpose: Perform at ring0 level. Shows how to operate at ring0 from visual basic several
	 ways, int 20,  calls and deviceiocontrol. Hopefully this program and the included
	 documentation will help you in some way.
	 
	
File List
	Class:
	asmdec.cls by '*****W32 OPCODE DISASSEMBLER WRITTEN BY VANJA FUCKAR EMAIL:INGA@VIP.HR
	http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=41582&lngWId=1
	purpose: translate machine level code in disassember mnemonic opcodes.

	Bas:
	modasmcalls.bas
	purpose: use hook into GDT to perform at ring0 in ASMGetCpu function
		 use int 20 and hook into GDT to perform at ring0 vmm service call
	modBas.bas 
	purpose declarations adn various machine language routines with differing parameters
	to perform ring0. Various registry procedures to obtain the static list of VxD's. Also
	memory copy routine.

	modCalls.bas
	purpose: just to show you all the various service calls of vmm to be either called
		or used with int 20
		ie:
		Call VMM_GetDDBList  - would be E8 xx xx xx xx  or whatever call you wish to use.
		or int 20            - would be CD 20
						3F 01 01 00    - the 01 00 lo - hi
								= 0001 which is the vmm service
								the 3F 01 lo -hi 
								= 013F which is GetDDBList
		all int 20 calls are handled the same way
		other notable windows services are
		vwin32, ifsmgr, shell, vmouse etc.

	moddecl.bas
	purpose: declarations file

	moduleASM
	purpose: used in conjunction with modasmcalls.cls

	Related Documents.
	This readme file
	GDT_info.txt a great article by Yann Stephen to give an explantion on how the
	protected mode works.

	vmm.inc - usefule information.
	
	assmTut_vxd.txt - The basic file used to create the blue screen of death Message.
			The vxd is accessed via deviceiocontrol. There is 1 service call
			and that is the blue screen that appears. from a call to 			                         SHELL_SYSMODAL_MESSAGE
	
       services.dat - list the major ring0 service calls
