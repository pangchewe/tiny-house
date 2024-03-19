<h2 align="center"> Web demo for "Tiny House Project"</h2>



## Author
pangchewe

## Description
<a href="https://pangchewe.github.io/tiny-house/" target="_blank"> THANIIE </a> is a website to showcase Tiny House Project. <!-- Built with love -->

## License
THANIIE is licensed under the **MIT License**.

=SUMIFS(ACF_Data!$H:$H,ACF_Data!$C:$C,"UA0102",ACF_Data!$K:$K,"GM000",ACF_Data!$A:$A,MONTH(AY$3)&"/"&YEAR(AY$3),ACF_Data!Z:Z, "AGENCY")


=SUMIFS('[Supplementary File Template_Jan2024_Enhancement_LLF_Pang.xlsx]ACF_Data'!$H:$H, '[Supplementary File Template_Jan2024_Enhancement_LLF_Pang.xlsx]ACF_Data'!$C:$C, "UA0102", '[Supplementary File Template_Jan2024_Enhancement_LLF_Pang.xlsx]ACF_Data'!$K:$K, "GM000", '[Supplementary File Template_Jan2024_Enhancement_LLF_Pang.xlsx]ACF_Data'!$A:$A, MONTH(AY$3)&"/"&YEAR(AY$3), '[Supplementary File Template_Jan2024_Enhancement_LLF_Pang.xlsx]ACF_Data'!Z:Z, "AGENCY")


=+(SUMIFS('GCS DataBase'!$G:$G,'GCS DataBase'!$H:$H,$C11,'GCS DataBase'!$B:$B,"GM000",'GCS DataBase'!$E:$E,YEAR(D$3)&TEXT(MONTH(D$3),"00"),'GCS DataBase'!$C:$C,"AGENCY"))*$AK11


=+(SUMIFS('[Supplementary File Template_Jan2024_Enhancement_LLF_Pang.xlsx]GCS DataBase'!$G:$G, '[Supplementary File Template_Jan2024_Enhancement_LLF_Pang.xlsx]GCS DataBase'!$H:$H, $C11, '[Supplementary File Template_Jan2024_Enhancement_LLF_Pang.xlsx]GCS DataBase'!$B:$B, "GM000", '[Supplementary File Template_Jan2024_Enhancement_LLF_Pang.xlsx]GCS DataBase'!$E:$E, YEAR(D$3)&TEXT(MONTH(D$3),"00"), '[Supplementary File Template_Jan2024_Enhancement_LLF_Pang.xlsx]GCS DataBase'!$C:$C, "AGENCY"))*$AK11

From
=-(SUMIFS(LLF_Accounting_Entries!$AG:$AG,LLF_Accounting_Entries!$C:$C,"6100000100",LLF_Accounting_Entries!$K:$K,"GM000",LLF_Accounting_Entries!$A:$A,'MTD Results Summary'!AY$3)-+SUMIFS(LLF_Accounting_Entries!$AG:$AG,LLF_Accounting_Entries!$C:$C,"6100000100",LLF_Accounting_Entries!$K:$K,"GM000",LLF_Accounting_Entries!$A:$A,'MTD Results Summary'!AX$3))
to
=-(SUMIFS('[Supplementary File Template_Jan2024_Enhancement_LLF_Pang.xlsx]LLF_Accounting_Entries'!$AG:$AG, '[Supplementary File Template_Jan2024_Enhancement_LLF_Pang.xlsx]LLF_Accounting_Entries'!$C:$C, "6100000100", '[Supplementary File Template_Jan2024_Enhancement_LLF_Pang.xlsx]LLF_Accounting_Entries'!$K:$K, "GM000", '[Supplementary File Template_Jan2024_Enhancement_LLF_Pang.xlsx]LLF_Accounting_Entries'!$A:$A, '[MTD Results Summary]MTD Results Summary'!AY$3, '[Supplementary File Template_Jan2024_Enhancement_LLF_Pang.xlsx]LLF_Accounting_Entries'!AH:AH, "AGENCY") - SUMIFS('[Supplementary File Template_Jan2024_Enhancement_LLF_Pang.xlsx]LLF_Accounting_Entries'!$AG:$AG, '[Supplementary File Template_Jan2024_Enhancement_LLF_Pang.xlsx]LLF_Accounting_Entries'!$C:$C, "6100000100", '[Supplementary File Template_Jan2024_Enhancement_LLF_Pang.xlsx]LLF_Accounting_Entries'!$K:$K, "GM000", '[Supplementary File Template_Jan2024_Enhancement_LLF_Pang.xlsx]LLF_Accounting_Entries'!$A:$A, '[MTD Results Summary]MTD Results Summary'!AX$3, '[Supplementary File Template_Jan2024_Enhancement_LLF_Pang.xlsx]LLF_Accounting_Entries'!AH:AH, "AGENCY"))
so make this 
=-(SUMIFS(LLF_Accounting_Entries!$AG:$AG,LLF_Accounting_Entries!$C:$C,"6100000100",LLF_Accounting_Entries!$K:$K,"GM000",LLF_Accounting_Entries!$A:$A,'MTD Results Summary'!AX$3)-+SUMIFS(LLF_Accounting_Entries!$AG:$AG,LLF_Accounting_Entries!$C:$C,"6100000100",LLF_Accounting_Entries!$K:$K,"GM000",LLF_Accounting_Entries!$A:$A,'MTD Results Summary'!AW$3))

1.
=-(SUMIFS(LLF_Accounting_Entries!$AG:$AG,LLF_Accounting_Entries!$C:$C,"6100000100",LLF_Accounting_Entries!$K:$K,"GM000",LLF_Accounting_Entries!$A:$A,'MTD Results Summary'!AW$3)-+SUMIFS(LLF_Accounting_Entries!$AG:$AG,LLF_Accounting_Entries!$C:$C,"6100000100",LLF_Accounting_Entries!$K:$K,"GM000",LLF_Accounting_Entries!$A:$A,'MTD Results Summary'!AV$3))
2.
=-(SUMIFS(LLF_Accounting_Entries!$AG:$AG,LLF_Accounting_Entries!$C:$C,"6100000100",LLF_Accounting_Entries!$K:$K,"GM000",LLF_Accounting_Entries!$A:$A,'MTD Results Summary'!AV$3)-+SUMIFS(LLF_Accounting_Entries!$AG:$AG,LLF_Accounting_Entries!$C:$C,"6100000100",LLF_Accounting_Entries!$K:$K,"GM000",LLF_Accounting_Entries!$A:$A,'MTD Results Summary'!AU$3))

3.=-(SUMIFS(LLF_Accounting_Entries!$AG:$AG,LLF_Accounting_Entries!$C:$C,"6100000100",LLF_Accounting_Entries!$K:$K,"GM000",LLF_Accounting_Entries!$A:$A,'MTD Results Summary'!AU$3)-+SUMIFS(LLF_Accounting_Entries!$AG:$AG,LLF_Accounting_Entries!$C:$C,"6100000100",LLF_Accounting_Entries!$K:$K,"GM000",LLF_Accounting_Entries!$A:$A,'MTD Results Summary'!AT$3))
4.
=-(SUMIFS(LLF_Accounting_Entries!$AG:$AG,LLF_Accounting_Entries!$C:$C,"6100000100",LLF_Accounting_Entries!$K:$K,"GM000",LLF_Accounting_Entries!$A:$A,'MTD Results Summary'!AT$3)-+SUMIFS(LLF_Accounting_Entries!$AG:$AG,LLF_Accounting_Entries!$C:$C,"6100000100",LLF_Accounting_Entries!$K:$K,"GM000",LLF_Accounting_Entries!$A:$A,'MTD Results Summary'!AS$3))
5.=-(SUMIFS(LLF_Accounting_Entries!$AG:$AG,LLF_Accounting_Entries!$C:$C,"6100000100",LLF_Accounting_Entries!$K:$K,"GM000",LLF_Accounting_Entries!$A:$A,'MTD Results Summary'!AS$3)-+SUMIFS(LLF_Accounting_Entries!$AG:$AG,LLF_Accounting_Entries!$C:$C,"6100000100",LLF_Accounting_Entries!$K:$K,"GM000",LLF_Accounting_Entries!$A:$A,'MTD Results Summary'!AR$3))
