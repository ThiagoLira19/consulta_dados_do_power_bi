def getPortPowerBIDesktop():
    import psutil
    import socket

    port = ""

    # Encontra o PID do Power BI (msmdsrv.exe)
    target_pid = None
    for p in psutil.process_iter(['pid', 'name', 'cmdline']):
        if 'msmdsrv.exe' in str(p.info['cmdline']):
            target_pid = p.info['pid']
            break

    if not target_pid:
        print("Power BI (msmdsrv.exe) não encontrado.")
    else:
        print(f"PID encontrado: {target_pid}")

        # Agora escaneia as portas usadas pelo processo
        for conn in psutil.net_connections(kind='tcp'):
            if conn.pid == target_pid and conn.status == 'LISTEN' and conn.laddr.ip == '127.0.0.1':
                print(f"Porta de conexão: {conn.laddr.port}")
                port = conn.laddr.port
                break

    return port