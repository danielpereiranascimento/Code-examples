*==================================================================
* Stock Verification and Alert Email
*==================================================================
* Verifica stock suficiente para encomendas de cliente (n_ndos=37)
* Envia email de alerta caso stock seja insuficiente
*==================================================================

Local mret, cTitIDU, cDir, Cliemail, email, utilizador, assinatura, cc

mret = .T.

If n_ndos = 37
    Select bi
    Scan
        If u_sqlexec([SELECT sa.stock, st.qttcli, sa.armazem 
                       FROM st (nolock) 
                       INNER JOIN sa ON st.ref = sa.ref 
                       WHERE sa.armazem=1 AND st.ref=] + bi.ref + ["], "temp")
            Select temp
            If temp.stock <> 0
                If bi.qtt + temp.qttcli > temp.stock
                    mret = .F.
                    Exit
                Endif
            Endif
        Endif
    Endscan

    * Preparar envio de email
    mostrameisto("temp", "xpto")
Endif

fecha("temp")
Return mret
