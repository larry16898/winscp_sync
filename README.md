# winscp_sync
refer to https://winscp.net/eng/docs/library_session_synchronizedirectories

# Obtaining host key fingerprint
ssh-keygen -l -f /etc/ssh/ssh_host_rsa_key

Since OpenSSH 6.8, you have to add the -E md5 switch to get the format needed for WinSCP.

ssh-keygen -l -f /etc/ssh/ssh_host_rsa_key -E md5
