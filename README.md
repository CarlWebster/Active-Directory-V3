# Active-Directory-V3
Active Directory V3 Documentation Script

    Creates a complete inventory of a Microsoft Active Directory Forest or Domain using
    Microsoft PowerShell, Word, plain text, or HTML.

    Creates a Word or PDF document, text, or HTML file named after the Active Directory
    Forest or Domain.

    Version 3.0 changes the default output report from Word to HTML.

    Word and PDF document includes a Cover Page, Table of Contents, and Footer.
    Includes support for the following language versions of Microsoft Word:
        Catalan
        Chinese
        Danish
        Dutch
        English
        Finnish
        French
        German
        Norwegian
        Portuguese
        Spanish
        Swedish

    The script requires at least PowerShell version 3 but runs best in version 5.

    Word is NOT needed to run the script. This script outputs in Text and HTML.

    You do NOT have to run this script on a domain controller. This script was developed
    and run from a Windows 10 VM.

    While most of the script can run with a non-admin account, there are some features
    that will not or may not work without domain admin or enterprise admin rights.
    The Hardware and Services parameters require domain admin privileges.

    The script does gathering of information on Time Server and AD database, log file, and
    SYSVOL locations. Those require access to the registry on each domain controller, which
    means the script should now always be run from an elevated PowerShell session with an
    account with a minimum of domain admin rights.

    Running the script in a forest with multiple domains requires Enterprise Admin rights.

    The count of all users may not be accurate if the user running the script does not have
    the necessary permissions on all user objects.  In that case, there may be user accounts
    classified as "unknown".

    To run the script from a workstation, RSAT is required.

    Remote Server Administration Tools for Windows 7 with Service Pack 1 (SP1)
        https://carlwebster.sharefile.com/d-sace5ee0f1ada47289ca14be544878a24

    Remote Server Administration Tools for Windows 8
        https://carlwebster.sharefile.com/d-s791075d451fc415ca83ec8958b385a95

    Remote Server Administration Tools for Windows 8.1
        https://carlwebster.sharefile.com/d-s37209afb73dc48f497745924ed854226

    Remote Server Administration Tools for Windows 10
        http://www.microsoft.com/en-us/download/details.aspx?id=45520
