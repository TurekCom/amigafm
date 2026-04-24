#![cfg_attr(target_os = "windows", windows_subsystem = "windows")]
#![allow(unsafe_op_in_unsafe_fn)]

mod nvda;

use std::collections::{HashMap, HashSet};
use std::ffi::{OsStr, c_void};
use std::fs;
use std::io;
use std::io::{Read, Write};
use std::net::{Ipv4Addr, SocketAddr, TcpStream, ToSocketAddrs};
use std::os::windows::ffi::OsStrExt;
use std::os::windows::fs::MetadataExt;
use std::os::windows::process::CommandExt;
use std::path::{Path, PathBuf};
use std::process::{Command, Output, Stdio};
use std::ptr::{null, null_mut};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::mpsc::{self, Receiver, Sender};
use std::sync::{Arc, Mutex};
use std::thread;
use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};

use mdns_sd::{ServiceDaemon, ServiceEvent};
use netdev::get_interfaces;
use netdev::interface::state::OperState;
use nfs3_client::nfs3_types::mount::exports;
use nfs3_client::nfs3_types::nfs3::{self, Nfs3Option, filename3};
use nfs3_client::nfs3_types::rpc::{auth_unix, opaque_auth};
use nfs3_client::nfs3_types::xdr_codec::Opaque;
use nfs3_client::tokio::{TokioConnector, TokioIo};
use nfs3_client::{MountClient, Nfs3ConnectionBuilder, PortmapperClient};
use nvda::NvdaController;
use regex::{Regex, RegexBuilder};
use remotefs::fs::{FileType as RemoteFileType, Metadata as RemoteMetadata, UnixPex};
use remotefs::{File as RemoteFile, RemoteFs};
use remotefs_ftp::FtpFs;
use remotefs_ssh::{NoCheckServerKey, RusshSession, SftpFs, SshKeyStorage, SshOpts};
use remotefs_webdav::WebDAVFs;
use serde::{Deserialize, Serialize};
use tokio::net::TcpStream as TokioTcpStream;
use windows_sys::Win32::Foundation::{
    CloseHandle, FILETIME, GetLastError, GlobalFree, HWND, LPARAM, LRESULT, LocalFree, RECT,
    SYSTEMTIME, WPARAM,
};
use windows_sys::Win32::Graphics::Gdi::{
    CLEARTYPE_QUALITY, CLIP_DEFAULT_PRECIS, COLOR_WINDOWFRAME, CreateFontW, CreateSolidBrush,
    DEFAULT_CHARSET, DEFAULT_PITCH, DeleteObject, FF_MODERN, FW_BOLD, GetStockObject, HBRUSH, HDC,
    HFONT, InvalidateRect, OUT_DEFAULT_PRECIS, SetBkColor, SetBkMode, SetTextColor, TRANSPARENT,
    UpdateWindow,
};
use windows_sys::Win32::NetworkManagement::NetManagement::{
    MAX_PREFERRED_LENGTH, NetApiBufferFree, NetServerEnum, SERVER_INFO_101, SV_TYPE_SERVER,
    SV_TYPE_SERVER_NT, SV_TYPE_SERVER_UNIX, SV_TYPE_WORKSTATION,
};
use windows_sys::Win32::Security::Cryptography::{
    CRYPT_INTEGER_BLOB, CRYPTPROTECT_UI_FORBIDDEN, CryptProtectData, CryptUnprotectData,
};
use windows_sys::Win32::Storage::FileSystem::{
    FILE_ATTRIBUTE_DIRECTORY, FileTimeToLocalFileTime, GetFileAttributesW, INVALID_FILE_ATTRIBUTES,
};
use windows_sys::Win32::System::Com::{
    CLSCTX_INPROC_SERVER, COINIT_APARTMENTTHREADED, COINIT_DISABLE_OLE1DDE, CoCreateInstance,
    CoInitializeEx, CoUninitialize, STGM_READ,
};
use windows_sys::Win32::System::DataExchange::{
    CloseClipboard, EmptyClipboard, GetClipboardData, GetClipboardSequenceNumber,
    IsClipboardFormatAvailable, OpenClipboard, RegisterClipboardFormatW, SetClipboardData,
};
use windows_sys::Win32::System::LibraryLoader::GetModuleHandleW;
use windows_sys::Win32::System::Memory::{
    GMEM_MOVEABLE, GMEM_ZEROINIT, GlobalAlloc, GlobalLock, GlobalUnlock,
};
use windows_sys::Win32::System::Threading::{GetExitCodeProcess, INFINITE, WaitForSingleObject};
use windows_sys::Win32::System::Time::FileTimeToSystemTime;
use windows_sys::Win32::UI::Controls::Dialogs::{
    GetOpenFileNameW, OFN_EXPLORER, OFN_FILEMUSTEXIST, OFN_HIDEREADONLY, OFN_NOCHANGEDIR,
    OFN_PATHMUSTEXIST, OPENFILENAMEW,
};
use windows_sys::Win32::UI::Controls::{
    EM_SETSEL, ICC_STANDARD_CLASSES, INITCOMMONCONTROLSEX, InitCommonControlsEx,
};
use windows_sys::Win32::UI::Input::KeyboardAndMouse::{
    EnableWindow, GetFocus, GetKeyState, SetFocus, VK_CONTROL, VK_SHIFT,
};
use windows_sys::Win32::UI::Shell::Common::ITEMIDLIST;
use windows_sys::Win32::UI::Shell::{
    CMF_CANRENAME, CMF_EXPLORE, CMF_NORMAL, CMINVOKECOMMANDINFO, DefSubclassProc, DragQueryFileW,
    ILFree, RemoveWindowSubclass, SEE_MASK_NOCLOSEPROCESS, SHBindToParent, SHELLEXECUTEINFOW,
    SHOP_FILEPATH, SHObjectProperties, SHParseDisplayName, SLGP_UNCPRIORITY,
    SetCurrentProcessExplicitAppUserModelID, SetWindowSubclass, ShellExecuteExW, ShellExecuteW,
    ShellLink,
};
use windows_sys::Win32::UI::WindowsAndMessaging::{
    AppendMenuW, BS_DEFPUSHBUTTON, BS_PUSHBUTTON, CB_ADDSTRING, CB_GETCURSEL, CB_SETCURSEL,
    CBN_SELCHANGE, CBS_DROPDOWNLIST, CS_HREDRAW, CS_VREDRAW, CW_USEDEFAULT, CreateMenu,
    CreatePopupMenu, CreateWindowExW, DefWindowProcW, DestroyMenu, DestroyWindow, DispatchMessageW,
    DrawMenuBar, GWLP_USERDATA, GetClientRect, GetDlgCtrlID, GetMenu, GetMessageW,
    GetWindowLongPtrW, GetWindowRect, GetWindowTextLengthW, GetWindowTextW, HMENU, IDC_ARROW,
    IDI_APPLICATION, IsDialogMessageW, LB_ADDSTRING, LB_GETCURSEL, LB_RESETCONTENT, LB_SETCURSEL,
    LB_SETTOPINDEX, LBN_DBLCLK, LBN_SELCHANGE, LBN_SETFOCUS, LBS_NOTIFY, LoadCursorW, LoadIconW,
    MB_ICONERROR, MB_OK, MF_CHECKED, MF_POPUP, MF_SEPARATOR, MF_STRING, MF_UNCHECKED, MSG,
    MessageBoxW, MoveWindow, PostMessageW, PostQuitMessage, RegisterClassExW, SW_MAXIMIZE, SW_SHOW,
    SendMessageW, SetMenu, SetWindowLongPtrW, SetWindowTextW, ShowWindow, TPM_LEFTALIGN,
    TPM_RETURNCMD, TPM_RIGHTBUTTON, TPM_TOPALIGN, TrackPopupMenu, TrackPopupMenuEx,
    TranslateMessage, WM_APP, WM_CLOSE, WM_COMMAND, WM_CREATE, WM_CTLCOLOREDIT, WM_CTLCOLORLISTBOX,
    WM_CTLCOLORSTATIC, WM_DESTROY, WM_GETDLGCODE, WM_KEYDOWN, WM_NCCREATE, WM_NCDESTROY,
    WM_SETFOCUS, WM_SIZE, WNDCLASSEXW, WS_BORDER, WS_CAPTION, WS_CHILD, WS_CLIPCHILDREN,
    WS_CLIPSIBLINGS, WS_EX_CLIENTEDGE, WS_EX_CONTROLPARENT, WS_EX_DLGMODALFRAME,
    WS_OVERLAPPEDWINDOW, WS_SYSMENU, WS_TABSTOP, WS_VISIBLE, WS_VSCROLL,
};
use windows_sys::core::GUID;

const IDC_LEFT_LABEL: i32 = 1001;
const IDC_RIGHT_LABEL: i32 = 1002;
const IDC_LEFT_LIST: i32 = 1101;
const IDC_RIGHT_LIST: i32 = 1102;
const IDC_STATUS: i32 = 1201;

const IDM_RENAME: u16 = 2001;
const IDM_COPY: u16 = 2002;
const IDM_MOVE: u16 = 2003;
const IDM_NEW_FOLDER: u16 = 2004;
const IDM_DELETE: u16 = 2005;
const IDM_REFRESH: u16 = 2006;
const IDM_QUIT: u16 = 2007;
const IDM_MARK_ALL: u16 = 2008;
const IDM_OPTIONS: u16 = 2009;
const IDM_ABOUT: u16 = 2010;
const IDM_PROPERTIES: u16 = 2011;
const IDM_PERMISSIONS: u16 = 2012;
const IDM_VIEW_SIZE: u16 = 2013;
const IDM_VIEW_TYPE: u16 = 2014;
const IDM_VIEW_CREATED: u16 = 2015;
const IDM_VIEW_MODIFIED: u16 = 2016;
const IDM_SYSTEM_CONTEXT: u16 = 2017;
const IDM_SEARCH: u16 = 2018;
const IDM_UNMARK_ALL: u16 = 2019;
const IDM_INVERT_MARKS: u16 = 2020;
const IDM_MARK_EXTENSION: u16 = 2021;
const IDM_MARK_NAME: u16 = 2022;
const IDM_ADD_TO_FAVORITES: u16 = 2023;
const IDM_ADD_NETWORK_CONNECTION: u16 = 2024;
const IDM_DISCOVER_NETWORK_SERVERS: u16 = 2025;
const IDM_EDIT_NETWORK_CONNECTION: u16 = 2026;
const IDM_REMOVE_NETWORK_CONNECTION: u16 = 2027;
const IDM_EXTRACT_HERE: u16 = 2028;
const IDM_EXTRACT_TO_FOLDER: u16 = 2029;
const IDM_EXTRACT_TO_OTHER_PANEL: u16 = 2030;
const IDM_CREATE_ARCHIVE: u16 = 2031;
const IDM_JOIN_SPLIT_ARCHIVE: u16 = 2032;
const IDM_CHECKSUM_CREATE: u16 = 2033;
const IDM_CHECKSUM_VERIFY: u16 = 2034;

const ID_DIALOG_EDIT: i32 = 3001;
const ID_DIALOG_OK: i32 = 3002;
const ID_DIALOG_CANCEL: i32 = 3003;
const ID_DIALOG_INFO: i32 = 3004;
const ID_DIALOG_YES: i32 = 3005;
const ID_DIALOG_NO: i32 = 3006;
const ID_DIALOG_REPLACE: i32 = 3007;
const ID_DIALOG_SKIP: i32 = 3008;
const ID_DIALOG_SEARCH_LOCAL: i32 = 3009;
const ID_DIALOG_SEARCH_RECURSIVE: i32 = 3010;
const ID_NET_PROTOCOL: i32 = 3011;
const ID_NET_HOST: i32 = 3017;
const ID_NET_USERNAME: i32 = 3018;
const ID_NET_PASSWORD: i32 = 3019;
const ID_NET_SSH_KEY: i32 = 3020;
const ID_NET_DIRECTORY: i32 = 3021;
const ID_NET_DISPLAY_NAME: i32 = 3022;
const ID_NET_ANONYMOUS: i32 = 3023;
const ID_NET_DISCOVERY_LIST: i32 = 3024;
const ID_NET_DISCOVERY_ADD: i32 = 3025;
const ID_NET_SSH_KEY_BROWSE: i32 = 3026;
const ID_ARCHIVE_TYPE: i32 = 3030;
const ID_ARCHIVE_NAME: i32 = 3031;
const ID_ARCHIVE_LEVEL: i32 = 3032;
const ID_ARCHIVE_ENCRYPTED: i32 = 3033;
const ID_ARCHIVE_ENCRYPTION: i32 = 3034;
const ID_ARCHIVE_PASSWORD: i32 = 3035;
const ID_ARCHIVE_VOLUME: i32 = 3036;

const WM_PANEL_ACTION: u32 = WM_APP + 1;
const WM_PANEL_SEARCH: u32 = WM_APP + 2;
const WM_PANEL_LOAD_EVENT: u32 = WM_APP + 3;
const WM_PROGRESS_EVENT: u32 = WM_APP + 4;
const WM_PROGRESS_NAVIGATE: u32 = WM_APP + 5;
const WM_DIALOG_NAVIGATE: u32 = WM_APP + 6;
const WM_SEARCH_EVENT: u32 = WM_APP + 7;
const WM_CUT_MSG: u32 = 0x0300;
const WM_COPY_MSG: u32 = 0x0301;
const WM_PASTE_MSG: u32 = 0x0302;

const MAIN_CLASS: &str = "AmigaFmNativeMain";
const INPUT_DIALOG_CLASS: &str = "AmigaFmNativeInputDialog";
const SEARCH_DIALOG_CLASS: &str = "AmigaFmNativeSearchDialog";
const OPERATION_PROMPT_DIALOG_CLASS: &str = "AmigaFmNativeOperationPrompt";
const PROGRESS_DIALOG_CLASS: &str = "AmigaFmNativeProgressDialog";
const NETWORK_DIALOG_CLASS: &str = "AmigaFmNativeNetworkDialog";
const DISCOVERY_DIALOG_CLASS: &str = "AmigaFmNativeDiscoveryDialog";
const ARCHIVE_CREATE_DIALOG_CLASS: &str = "AmigaFmNativeArchiveCreateDialog";

const BLACK: u32 = rgb(0, 0, 0);
const YELLOW: u32 = rgb(255, 220, 0);
const PANEL_LOAD_CHUNK_SIZE: usize = 500;
const DISCOVERY_CONCURRENCY: usize = 24;
const DISCOVERY_CACHE_TTL: Duration = Duration::from_secs(300);
const DISCOVERY_MAX_HOSTS_PER_INTERFACE: usize = 254;
const DISCOVERY_CONNECT_TIMEOUT: Duration = Duration::from_millis(850);
const DISCOVERY_IO_TIMEOUT: Duration = Duration::from_millis(900);
const MDNS_DISCOVERY_TIMEOUT: Duration = Duration::from_secs(3);
const SMB_SHARE_ENUM_TIMEOUT: Duration = Duration::from_secs(2);
const DISCOVERY_MIN_RESULTS_BEFORE_FALLBACK: usize = 2;
const BS_AUTOCHECKBOX: u32 = 0x00000003;
const ES_PASSWORD: u32 = 0x00000020;
const BM_GETCHECK: u32 = 0x00F0;
const BM_SETCHECK: u32 = 0x00F1;
const BST_CHECKED: usize = 1;
const WS_GROUP_STYLE: u32 = 0x00020000;
const CF_HDROP_FORMAT: u32 = 15;
const DROPEFFECT_COPY_VALUE: u32 = 1;
const DROPEFFECT_MOVE_VALUE: u32 = 2;
const CREATE_NO_WINDOW_FLAG: u32 = 0x08000000;

const fn rgb(r: u8, g: u8, b: u8) -> u32 {
    (r as u32) | ((g as u32) << 8) | ((b as u32) << 16)
}

#[derive(Clone, Copy, Default, Serialize, Deserialize)]
#[serde(default)]
struct ViewOptions {
    show_size: bool,
    show_type: bool,
    show_created: bool,
    show_modified: bool,
}

impl ViewOptions {
    const fn requires_metadata(self) -> bool {
        self.show_size || self.show_created || self.show_modified
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
enum PanelLocation {
    Filesystem(PathBuf),
    Remote(RemoteLocation),
    Archive(ArchiveLocation),
    Drives,
    NetworkResources,
    FavoriteDirectories,
    FavoriteFiles,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ClipboardOperation {
    Copy,
    Move,
}

#[derive(Clone)]
struct AppClipboardState {
    operation: ClipboardOperation,
    source_panel_index: Option<usize>,
    source_remote: Option<RemoteLocation>,
    source_entries: Vec<PanelEntry>,
    items: Vec<PathBuf>,
    staged_paths: Vec<PathBuf>,
    clipboard_sequence: u32,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum EntryKind {
    Loading,
    SearchLoading,
    SearchEmpty,
    GoToDrives,
    NetworkResource,
    FavoriteDirectoriesRoot,
    FavoriteFilesRoot,
    Parent,
    Directory,
    File,
    Drive,
    NetworkPlaceholder,
}

#[derive(Clone)]
struct PanelEntry {
    name: String,
    path: Option<PathBuf>,
    link_target: Option<PathBuf>,
    network_resource: Option<NetworkResource>,
    kind: EntryKind,
    size_bytes: Option<u64>,
    type_label: Option<String>,
    created_label: Option<String>,
    modified_label: Option<String>,
}

#[derive(Clone)]
struct SearchState {
    pattern: String,
    recursive: bool,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct ArchiveLocation {
    archive_path: PathBuf,
    inside_path: PathBuf,
}

impl ArchiveLocation {
    fn root(archive_path: PathBuf) -> Self {
        Self {
            archive_path,
            inside_path: PathBuf::new(),
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct RemoteLocation {
    resource: NetworkResource,
    original_resource: NetworkResource,
    path: PathBuf,
    sftp_privilege_mode: SftpPrivilegeMode,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
enum SftpPrivilegeMode {
    #[default]
    Normal,
    Root,
    Sudo,
}

impl RemoteLocation {
    fn uses_sftp_sudo(&self) -> bool {
        self.resource.protocol == NetworkProtocol::Sftp
            && self.sftp_privilege_mode == SftpPrivilegeMode::Sudo
    }

    fn uses_sftp_root(&self) -> bool {
        self.resource.protocol == NetworkProtocol::Sftp
            && self.sftp_privilege_mode == SftpPrivilegeMode::Root
    }

    fn downgraded_sftp_location(&self, path: PathBuf) -> RemoteLocation {
        RemoteLocation {
            resource: self.original_resource.clone(),
            original_resource: self.original_resource.clone(),
            path,
            sftp_privilege_mode: SftpPrivilegeMode::Normal,
        }
    }
}

#[derive(Clone, Default)]
struct AppSettings {
    view_options: ViewOptions,
    favorite_directories: Vec<PathBuf>,
    favorite_files: Vec<PathBuf>,
    network_resources: Vec<NetworkResource>,
    discovery_cache: DiscoveryCache,
}

#[derive(Default, Serialize, Deserialize)]
#[serde(default)]
struct PersistedAppSettings {
    view_options: ViewOptions,
    favorite_directories: Vec<PathBuf>,
    favorite_files: Vec<PathBuf>,
    network_resources: Vec<PersistedNetworkResource>,
    discovery_cache: Vec<PersistedDiscoveryCacheEntry>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
enum NetworkProtocol {
    Sftp,
    Ftp,
    Ftps,
    Nfs,
    WebDav,
    Smb,
}

impl Default for NetworkProtocol {
    fn default() -> Self {
        Self::Smb
    }
}

impl NetworkProtocol {
    const ALL: [Self; 6] = [
        Self::Sftp,
        Self::Ftp,
        Self::Ftps,
        Self::Nfs,
        Self::WebDav,
        Self::Smb,
    ];

    const fn label(self) -> &'static str {
        match self {
            Self::Sftp => "SFTP",
            Self::Ftp => "FTP",
            Self::Ftps => "FTPS",
            Self::Nfs => "NFS",
            Self::WebDav => "WebDAV",
            Self::Smb => "SMB",
        }
    }

    const fn inferred_type_label(self) -> &'static str {
        match self {
            Self::Sftp => "zasób SFTP",
            Self::Ftp => "zasób FTP",
            Self::Ftps => "zasób FTPS",
            Self::Nfs => "zasób NFS",
            Self::WebDav => "zasób WebDAV",
            Self::Smb => "zasób SMB",
        }
    }
}

#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(default)]
struct NetworkResource {
    protocol: NetworkProtocol,
    host: String,
    username: String,
    password: String,
    root_password: String,
    sudo_password: String,
    ssh_key: String,
    default_directory: String,
    display_name: String,
    anonymous: bool,
}

struct FixedSshKeyStorage {
    path: PathBuf,
}

impl SshKeyStorage for FixedSshKeyStorage {
    fn resolve(&self, _host: &str, _username: &str) -> Option<PathBuf> {
        Some(self.path.clone())
    }
}

enum RemoteClient {
    Ftp(FtpFs),
    Sftp(SftpFs<RusshSession<NoCheckServerKey>>),
    WebDav(WebDAVFs),
    Nfs(NfsSession),
}

type NfsConnection = nfs3_client::Nfs3Connection<TokioIo<TokioTcpStream>>;

struct NfsSession {
    runtime: Arc<tokio::runtime::Runtime>,
    connection: Option<NfsConnection>,
    resource: NetworkResource,
}

impl NfsSession {
    fn new(resource: NetworkResource) -> io::Result<Self> {
        let runtime = Arc::new(
            tokio::runtime::Builder::new_current_thread()
                .enable_all()
                .build()
                .map_err(|error| io::Error::other(error.to_string()))?,
        );
        let mut session = Self {
            runtime,
            connection: None,
            resource,
        };
        session.connect()?;
        Ok(session)
    }

    fn connect(&mut self) -> io::Result<()> {
        if self.connection.is_some() {
            return Ok(());
        }
        let (host, portmapper_port) = parse_host_and_port(&self.resource.host, 111);
        if host.is_empty() {
            return Err(io::Error::other("brak hosta połączenia"));
        }
        let resolved_host = resolve_nfs_host(&host)?;
        let requested_mount_path = normalize_nfs_mount_path(&self.resource.default_directory);
        let exports = list_nfs_exports(&resolved_host, portmapper_port).unwrap_or_default();
        let mount_path = if !self.resource.default_directory.trim().is_empty() {
            exports
                .iter()
                .find(|export| export.eq_ignore_ascii_case(&requested_mount_path))
                .cloned()
                .unwrap_or(requested_mount_path)
        } else if exports.iter().any(|export| export == "/") {
            "/".to_string()
        } else if exports.len() == 1 {
            exports[0].clone()
        } else if !exports.is_empty() {
            return Err(io::Error::other(format!(
                "podaj eksport NFS, dostępne eksporty: {}",
                format_nfs_exports(&exports)
            )));
        } else {
            requested_mount_path
        };
        let credential = nfs_auth_credential();
        let connection = self
            .runtime
            .block_on(async {
                Nfs3ConnectionBuilder::new(TokioConnector, resolved_host, mount_path)
                    .portmapper_port(portmapper_port)
                    .credential(credential)
                    .mount()
                    .await
            })
            .map_err(nfs_error_to_io)?;
        self.connection = Some(connection);
        Ok(())
    }

    fn disconnect(&mut self) {
        if let Some(connection) = self.connection.take() {
            let _ = self
                .runtime
                .block_on(connection.unmount())
                .map_err(nfs_error_to_io);
        }
    }

    fn list_dir(&mut self, path: &Path) -> io::Result<Vec<RemoteFile>> {
        self.connect()?;
        let connection = self
            .connection
            .as_mut()
            .ok_or_else(|| io::Error::other("brak połączenia NFS"))?;
        self.runtime.block_on(nfs_list_dir(connection, path))
    }

    fn stat(&mut self, path: &Path) -> io::Result<RemoteFile> {
        self.connect()?;
        let connection = self
            .connection
            .as_mut()
            .ok_or_else(|| io::Error::other("brak połączenia NFS"))?;
        self.runtime.block_on(nfs_stat(connection, path))
    }

    fn exists(&mut self, path: &Path) -> io::Result<bool> {
        self.connect()?;
        let connection = self
            .connection
            .as_mut()
            .ok_or_else(|| io::Error::other("brak połączenia NFS"))?;
        self.runtime.block_on(nfs_exists(connection, path))
    }

    fn create_dir(&mut self, path: &Path) -> io::Result<()> {
        self.connect()?;
        let connection = self
            .connection
            .as_mut()
            .ok_or_else(|| io::Error::other("brak połączenia NFS"))?;
        self.runtime.block_on(nfs_create_dir(connection, path))
    }

    fn rename(&mut self, src: &Path, dest: &Path) -> io::Result<()> {
        self.connect()?;
        let connection = self
            .connection
            .as_mut()
            .ok_or_else(|| io::Error::other("brak połączenia NFS"))?;
        self.runtime.block_on(nfs_rename(connection, src, dest))
    }

    fn remove_file(&mut self, path: &Path) -> io::Result<()> {
        self.connect()?;
        let connection = self
            .connection
            .as_mut()
            .ok_or_else(|| io::Error::other("brak połączenia NFS"))?;
        self.runtime.block_on(nfs_remove_file(connection, path))
    }

    fn remove_dir(&mut self, path: &Path) -> io::Result<()> {
        self.connect()?;
        let connection = self
            .connection
            .as_mut()
            .ok_or_else(|| io::Error::other("brak połączenia NFS"))?;
        self.runtime.block_on(nfs_remove_dir(connection, path))
    }

    fn download_to(&mut self, src: &Path, mut dest: Box<dyn Write + Send>) -> io::Result<u64> {
        self.connect()?;
        let connection = self
            .connection
            .as_mut()
            .ok_or_else(|| io::Error::other("brak połączenia NFS"))?;
        self.runtime
            .block_on(nfs_download_file(connection, src, &mut dest, None))
    }

    fn create_file(
        &mut self,
        path: &Path,
        metadata: &RemoteMetadata,
        mut reader: Box<dyn Read + Send>,
    ) -> io::Result<u64> {
        self.connect()?;
        let connection = self
            .connection
            .as_mut()
            .ok_or_else(|| io::Error::other("brak połączenia NFS"))?;
        self.runtime.block_on(nfs_upload_file(
            connection,
            path,
            metadata.size,
            &mut reader,
            None,
        ))
    }
}

#[derive(Clone, Debug, Default, Serialize, Deserialize)]
#[serde(default)]
struct PersistedNetworkResource {
    protocol: NetworkProtocol,
    host: String,
    username: String,
    encrypted_password: String,
    encrypted_root_password: String,
    encrypted_sudo_password: String,
    encrypted_ssh_key: String,
    password: String,
    root_password: String,
    sudo_password: String,
    ssh_key: String,
    default_directory: String,
    display_name: String,
    anonymous: bool,
}

impl NetworkResource {
    fn normalized_host(&self) -> String {
        self.host.trim().to_string()
    }

    fn effective_display_name(&self) -> String {
        let display_name = self.display_name.trim();
        if !display_name.is_empty() {
            display_name.to_string()
        } else {
            self.normalized_host()
        }
    }

    fn stable_key(&self) -> String {
        format!(
            "{}:{}:{}",
            self.protocol.label(),
            self.normalized_host().to_lowercase(),
            self.default_directory.trim().to_lowercase()
        )
    }

    fn summary_line(&self) -> String {
        let display_name = self.display_name.trim();
        let auth =
            if self.anonymous || (self.username.trim().is_empty() && self.password.is_empty()) {
                "bez danych logowania".to_string()
            } else if self.username.trim().is_empty() {
                "z hasłem".to_string()
            } else {
                format!("użytkownik {}", self.username.trim())
            };
        let directory = if self.default_directory.trim().is_empty() {
            "katalog domyślny nieustawiony".to_string()
        } else {
            format!("katalog {}", self.default_directory.trim())
        };
        if !display_name.is_empty() && !display_name.eq_ignore_ascii_case(&self.normalized_host()) {
            format!(
                "{}, {} {}, {}, {}",
                display_name,
                self.protocol.label(),
                self.normalized_host(),
                auth,
                directory
            )
        } else {
            format!(
                "{} {}, {}, {}",
                self.protocol.label(),
                self.normalized_host(),
                auth,
                directory
            )
        }
    }

    fn launch_target(&self) -> String {
        let host = self.normalized_host();
        let directory = self.default_directory.trim().trim_matches('/');
        match self.protocol {
            NetworkProtocol::Smb => {
                if directory.is_empty() {
                    format!("\\\\{host}")
                } else {
                    format!("\\\\{host}\\{}", directory.replace('/', "\\"))
                }
            }
            NetworkProtocol::Nfs => {
                if directory.is_empty() {
                    format!("nfs://{host}")
                } else {
                    format!("nfs://{host}/{directory}")
                }
            }
            NetworkProtocol::Sftp => build_network_uri("sftp", &host, directory, self),
            NetworkProtocol::Ftp => build_network_uri("ftp", &host, directory, self),
            NetworkProtocol::Ftps => build_network_uri("ftps", &host, directory, self),
            NetworkProtocol::WebDav => {
                let scheme = if host.starts_with("https://") {
                    "https"
                } else {
                    "http"
                };
                let clean_host = host
                    .trim_start_matches("http://")
                    .trim_start_matches("https://")
                    .to_string();
                build_network_uri(scheme, &clean_host, directory, self)
            }
        }
    }

    fn remote_root(&self) -> PathBuf {
        normalize_remote_directory(&self.default_directory)
    }

    fn as_remote_location(&self) -> Option<RemoteLocation> {
        if matches!(
            self.protocol,
            NetworkProtocol::Sftp
                | NetworkProtocol::Ftp
                | NetworkProtocol::Ftps
                | NetworkProtocol::Nfs
                | NetworkProtocol::WebDav
        ) {
            Some(RemoteLocation {
                resource: self.clone(),
                original_resource: self.clone(),
                path: if self.protocol == NetworkProtocol::Nfs {
                    PathBuf::from("/")
                } else {
                    self.remote_root()
                },
                sftp_privilege_mode: SftpPrivilegeMode::Normal,
            })
        } else {
            None
        }
    }

    fn smb_filesystem_path(&self) -> Option<PathBuf> {
        if self.protocol != NetworkProtocol::Smb {
            return None;
        }
        let host = self.normalized_host();
        if host.is_empty() {
            return None;
        }
        let directory = self.default_directory.trim().trim_matches('/');
        if directory.is_empty() {
            None
        } else {
            Some(PathBuf::from(format!(
                "\\\\{}\\{}",
                host,
                directory.replace('/', "\\")
            )))
        }
    }
}

impl PersistedNetworkResource {
    fn from_runtime(resource: &NetworkResource) -> io::Result<Self> {
        Ok(Self {
            protocol: resource.protocol,
            host: resource.host.clone(),
            username: resource.username.clone(),
            encrypted_password: protect_secret(&resource.password)?,
            encrypted_root_password: protect_secret(&resource.root_password)?,
            encrypted_sudo_password: protect_secret(&resource.sudo_password)?,
            encrypted_ssh_key: protect_secret(&resource.ssh_key)?,
            password: String::new(),
            root_password: String::new(),
            sudo_password: String::new(),
            ssh_key: String::new(),
            default_directory: resource.default_directory.clone(),
            display_name: resource.display_name.clone(),
            anonymous: resource.anonymous,
        })
    }

    fn into_runtime(self) -> NetworkResource {
        NetworkResource {
            protocol: self.protocol,
            host: self.host,
            username: self.username,
            password: unprotect_secret(&self.encrypted_password).unwrap_or(self.password),
            root_password: unprotect_secret(&self.encrypted_root_password)
                .unwrap_or(self.root_password),
            sudo_password: unprotect_secret(&self.encrypted_sudo_password)
                .unwrap_or(self.sudo_password),
            ssh_key: unprotect_secret(&self.encrypted_ssh_key).unwrap_or(self.ssh_key),
            default_directory: self.default_directory,
            display_name: self.display_name,
            anonymous: self.anonymous,
        }
    }
}

impl RemoteClient {
    fn connect(&mut self) -> io::Result<()> {
        match self {
            Self::Ftp(client) => client.connect().map(|_| ()).map_err(remote_error_to_io),
            Self::Sftp(client) => client.connect().map(|_| ()).map_err(remote_error_to_io),
            Self::WebDav(client) => client.connect().map(|_| ()).map_err(remote_error_to_io),
            Self::Nfs(client) => client.connect(),
        }
    }

    fn disconnect(&mut self) {
        let _ = match self {
            Self::Ftp(client) => client.disconnect(),
            Self::Sftp(client) => client.disconnect(),
            Self::WebDav(client) => client.disconnect(),
            Self::Nfs(client) => {
                client.disconnect();
                Ok(())
            }
        };
    }

    fn list_dir(&mut self, path: &Path) -> io::Result<Vec<RemoteFile>> {
        match self {
            Self::Ftp(client) => client.list_dir(path).map_err(remote_error_to_io),
            Self::Sftp(client) => client.list_dir(path).map_err(remote_error_to_io),
            Self::WebDav(client) => client.list_dir(path).map_err(remote_error_to_io),
            Self::Nfs(client) => client.list_dir(path),
        }
    }

    fn stat(&mut self, path: &Path) -> io::Result<RemoteFile> {
        match self {
            Self::Ftp(client) => client.stat(path).map_err(remote_error_to_io),
            Self::Sftp(client) => client.stat(path).map_err(remote_error_to_io),
            Self::WebDav(client) => client.stat(path).map_err(remote_error_to_io),
            Self::Nfs(client) => client.stat(path),
        }
    }

    fn exists(&mut self, path: &Path) -> io::Result<bool> {
        match self {
            Self::Ftp(client) => client.exists(path).map_err(remote_error_to_io),
            Self::Sftp(client) => client.exists(path).map_err(remote_error_to_io),
            Self::WebDav(client) => client.exists(path).map_err(remote_error_to_io),
            Self::Nfs(client) => client.exists(path),
        }
    }

    fn create_dir(&mut self, path: &Path) -> io::Result<()> {
        match self {
            Self::Ftp(client) => client
                .create_dir(path, UnixPex::from(0o755))
                .map_err(remote_error_to_io),
            Self::Sftp(client) => client
                .create_dir(path, UnixPex::from(0o755))
                .map_err(remote_error_to_io),
            Self::WebDav(client) => client
                .create_dir(path, UnixPex::from(0o755))
                .map_err(remote_error_to_io),
            Self::Nfs(client) => client.create_dir(path),
        }
    }

    fn rename(&mut self, src: &Path, dest: &Path) -> io::Result<()> {
        match self {
            Self::Ftp(client) => client.mov(src, dest).map_err(remote_error_to_io),
            Self::Sftp(client) => client.mov(src, dest).map_err(remote_error_to_io),
            Self::WebDav(client) => client.mov(src, dest).map_err(remote_error_to_io),
            Self::Nfs(client) => client.rename(src, dest),
        }
    }

    fn remove_file(&mut self, path: &Path) -> io::Result<()> {
        match self {
            Self::Ftp(client) => client.remove_file(path).map_err(remote_error_to_io),
            Self::Sftp(client) => client.remove_file(path).map_err(remote_error_to_io),
            Self::WebDav(client) => client.remove_file(path).map_err(remote_error_to_io),
            Self::Nfs(client) => client.remove_file(path),
        }
    }

    fn remove_dir(&mut self, path: &Path) -> io::Result<()> {
        match self {
            Self::Ftp(client) => client.remove_dir(path).map_err(remote_error_to_io),
            Self::Sftp(client) => client.remove_dir(path).map_err(remote_error_to_io),
            Self::WebDav(client) => client.remove_dir(path).map_err(remote_error_to_io),
            Self::Nfs(client) => client.remove_dir(path),
        }
    }

    fn download_to(&mut self, src: &Path, dest: Box<dyn Write + Send>) -> io::Result<u64> {
        match self {
            Self::Ftp(client) => client.open_file(src, dest).map_err(remote_error_to_io),
            Self::Sftp(client) => client.open_file(src, dest).map_err(remote_error_to_io),
            Self::WebDav(client) => client.open_file(src, dest).map_err(remote_error_to_io),
            Self::Nfs(client) => client.download_to(src, dest),
        }
    }

    fn create_file(
        &mut self,
        path: &Path,
        metadata: &RemoteMetadata,
        reader: Box<dyn Read + Send>,
    ) -> io::Result<u64> {
        match self {
            Self::Ftp(client) => client
                .create_file(path, metadata, reader)
                .map_err(remote_error_to_io),
            Self::Sftp(client) => client
                .create_file(path, metadata, reader)
                .map_err(remote_error_to_io),
            Self::WebDav(client) => client
                .create_file(path, metadata, reader)
                .map_err(remote_error_to_io),
            Self::Nfs(client) => client.create_file(path, metadata, reader),
        }
    }

    fn exec(&mut self, cmd: &str) -> io::Result<(u32, String)> {
        match self {
            Self::Sftp(client) => client.exec(cmd).map_err(remote_error_to_io),
            _ => Err(io::Error::other(
                "wykonywanie poleceń jest dostępne tylko dla SFTP",
            )),
        }
    }
}

#[derive(Clone, Debug, Serialize, Deserialize)]
struct DiscoveredServer {
    host: String,
    resolved_name: Option<String>,
    protocol: NetworkProtocol,
    default_directory: String,
    detail: Option<String>,
}

#[derive(Clone)]
struct DiscoveryCachedHost {
    scanned_at_epoch_secs: u64,
    services: Vec<DiscoveredServer>,
}

#[derive(Clone, Default)]
struct DiscoveryCache {
    hosts: HashMap<String, DiscoveryCachedHost>,
}

#[derive(Clone, Serialize, Deserialize)]
struct PersistedDiscoveryCacheEntry {
    host: String,
    scanned_at_epoch_secs: u64,
    services: Vec<DiscoveredServer>,
}

#[derive(Clone)]
struct DiscoveryScanResult {
    servers: Vec<DiscoveredServer>,
    probed_hosts: Vec<(String, Vec<DiscoveredServer>)>,
    cache_hits: usize,
    processed_hosts: usize,
    total_hosts: usize,
    elapsed: Duration,
}

#[derive(Clone, Copy)]
struct MdnsBrowseSpec {
    service_type: &'static str,
    protocol: NetworkProtocol,
    detail: &'static str,
}

impl DiscoveryCache {
    fn lookup_fresh(&self, host: &str) -> Option<&DiscoveryCachedHost> {
        let key = host.trim().to_ascii_lowercase();
        let cached = self.hosts.get(&key)?;
        let age = current_unix_timestamp().saturating_sub(cached.scanned_at_epoch_secs);
        if age <= DISCOVERY_CACHE_TTL.as_secs() {
            Some(cached)
        } else {
            None
        }
    }

    fn update_from_scan(&mut self, scan: &DiscoveryScanResult) {
        let now = current_unix_timestamp();
        for (host, services) in &scan.probed_hosts {
            self.hosts.insert(
                host.trim().to_ascii_lowercase(),
                DiscoveryCachedHost {
                    scanned_at_epoch_secs: now,
                    services: services.clone(),
                },
            );
        }
    }

    fn from_persisted(entries: Vec<PersistedDiscoveryCacheEntry>) -> Self {
        let mut hosts = HashMap::new();
        let now = current_unix_timestamp();
        for entry in entries {
            if now.saturating_sub(entry.scanned_at_epoch_secs) > DISCOVERY_CACHE_TTL.as_secs() {
                continue;
            }
            hosts.insert(
                entry.host.trim().to_ascii_lowercase(),
                DiscoveryCachedHost {
                    scanned_at_epoch_secs: entry.scanned_at_epoch_secs,
                    services: entry.services,
                },
            );
        }
        Self { hosts }
    }

    fn to_persisted(&self) -> Vec<PersistedDiscoveryCacheEntry> {
        let now = current_unix_timestamp();
        let mut entries = self
            .hosts
            .iter()
            .filter_map(|(host, cached)| {
                if now.saturating_sub(cached.scanned_at_epoch_secs) > DISCOVERY_CACHE_TTL.as_secs()
                {
                    None
                } else {
                    Some(PersistedDiscoveryCacheEntry {
                        host: host.clone(),
                        scanned_at_epoch_secs: cached.scanned_at_epoch_secs,
                        services: cached.services.clone(),
                    })
                }
            })
            .collect::<Vec<_>>();
        entries.sort_by(|left, right| left.host.cmp(&right.host));
        entries
    }
}

impl DiscoveredServer {
    fn display_line(&self) -> String {
        let mut details = vec![self.protocol.label().to_string()];
        if let Some(name) = &self.resolved_name {
            if !name.trim().is_empty() && !name.eq_ignore_ascii_case(&self.host) {
                details.insert(0, format!("nazwa {}", name.trim()));
            }
        }
        if self.protocol == NetworkProtocol::Nfs && !self.default_directory.trim().is_empty() {
            details.push(format!("eksport {}", self.default_directory.trim()));
        }
        if let Some(detail) = &self.detail {
            if !detail.trim().is_empty() {
                details.push(detail.trim().to_string());
            }
        }
        format!("{} ({})", self.host, details.join(", "))
    }

    fn to_resource_template(&self) -> NetworkResource {
        let base_name = self
            .resolved_name
            .as_ref()
            .map(|name| name.trim())
            .filter(|name| !name.is_empty())
            .unwrap_or(&self.host);
        let display_name =
            if self.protocol == NetworkProtocol::Nfs && !self.default_directory.trim().is_empty() {
                format!("{} {}", base_name, self.default_directory.trim())
            } else {
                base_name.to_string()
            };
        NetworkResource {
            protocol: self.protocol,
            host: self.host.clone(),
            default_directory: self.default_directory.clone(),
            display_name,
            anonymous: true,
            ..NetworkResource::default()
        }
    }
}

impl PanelEntry {
    fn loading() -> Self {
        Self {
            name: "Ładowanie katalogu...".to_string(),
            path: None,
            link_target: None,
            network_resource: None,
            kind: EntryKind::Loading,
            size_bytes: None,
            type_label: None,
            created_label: None,
            modified_label: None,
        }
    }

    fn search_loading() -> Self {
        Self {
            name: "Trwa wyszukiwanie".to_string(),
            path: None,
            link_target: None,
            network_resource: None,
            kind: EntryKind::SearchLoading,
            size_bytes: None,
            type_label: None,
            created_label: None,
            modified_label: None,
        }
    }

    fn search_empty() -> Self {
        Self {
            name: "Brak wyników".to_string(),
            path: None,
            link_target: None,
            network_resource: None,
            kind: EntryKind::SearchEmpty,
            size_bytes: None,
            type_label: None,
            created_label: None,
            modified_label: None,
        }
    }

    fn go_to_drives() -> Self {
        Self {
            name: "Przejdź do dysków".to_string(),
            path: None,
            link_target: None,
            network_resource: None,
            kind: EntryKind::GoToDrives,
            size_bytes: None,
            type_label: None,
            created_label: None,
            modified_label: None,
        }
    }

    fn favorite_directories_root() -> Self {
        Self {
            name: "Ulubione katalogi".to_string(),
            path: None,
            link_target: None,
            network_resource: None,
            kind: EntryKind::FavoriteDirectoriesRoot,
            size_bytes: None,
            type_label: Some("ulubione katalogi".to_string()),
            created_label: None,
            modified_label: None,
        }
    }

    fn favorite_files_root() -> Self {
        Self {
            name: "Ulubione pliki".to_string(),
            path: None,
            link_target: None,
            network_resource: None,
            kind: EntryKind::FavoriteFilesRoot,
            size_bytes: None,
            type_label: Some("ulubione pliki".to_string()),
            created_label: None,
            modified_label: None,
        }
    }

    fn parent(path: PathBuf) -> Self {
        Self {
            name: "..".to_string(),
            path: Some(path),
            link_target: None,
            network_resource: None,
            kind: EntryKind::Parent,
            size_bytes: None,
            type_label: None,
            created_label: None,
            modified_label: None,
        }
    }

    fn drive(path: PathBuf) -> Self {
        Self {
            name: path.display().to_string(),
            path: Some(path),
            link_target: None,
            network_resource: None,
            kind: EntryKind::Drive,
            size_bytes: None,
            type_label: Some("dysk".to_string()),
            created_label: None,
            modified_label: None,
        }
    }

    fn network_placeholder() -> Self {
        Self {
            name: "Zasoby sieciowe".to_string(),
            path: None,
            link_target: None,
            network_resource: None,
            kind: EntryKind::NetworkPlaceholder,
            size_bytes: None,
            type_label: Some("zasób sieciowy".to_string()),
            created_label: None,
            modified_label: None,
        }
    }

    fn from_remote_file(file: RemoteFile, view_options: ViewOptions) -> Self {
        let metadata = file.metadata().clone();
        let link_target = metadata
            .symlink
            .as_ref()
            .filter(|_| metadata.is_symlink())
            .map(|target| resolve_remote_symlink_path(file.path(), target));
        let kind = if file.is_dir() {
            EntryKind::Directory
        } else {
            EntryKind::File
        };
        let name = file.name();
        Self {
            name: name.clone(),
            path: Some(file.path().to_path_buf()),
            link_target,
            network_resource: None,
            kind,
            size_bytes: if view_options.show_size && matches!(kind, EntryKind::File) {
                Some(metadata.size)
            } else {
                None
            },
            type_label: if metadata.is_symlink() {
                Some("link symboliczny".to_string())
            } else {
                match kind {
                    EntryKind::Directory => Some("katalog".to_string()),
                    EntryKind::File => Some(infer_file_type(&name)),
                    _ => None,
                }
            },
            created_label: if view_options.show_created {
                metadata.created.and_then(format_system_time_label)
            } else {
                None
            },
            modified_label: if view_options.show_modified {
                metadata.modified.and_then(format_system_time_label)
            } else {
                None
            },
        }
    }

    fn from_dir_entry(entry: fs::DirEntry, view_options: ViewOptions) -> io::Result<Self> {
        Self::from_path(entry.path(), view_options)
    }

    fn from_path(path: PathBuf, view_options: ViewOptions) -> io::Result<Self> {
        let metadata = fs::symlink_metadata(&path)?;
        let file_type = metadata.file_type();
        let link_target = if file_type.is_symlink() {
            fs::read_link(&path)
                .ok()
                .map(|target| resolve_local_symlink_path(&path, &target))
        } else {
            None
        };
        let shortcut = is_windows_shortcut_path(&path);
        let followed_metadata = if file_type.is_symlink() {
            fs::metadata(&path).ok()
        } else {
            None
        };
        let kind = if file_type.is_dir()
            || followed_metadata
                .as_ref()
                .map(|metadata| metadata.is_dir())
                .unwrap_or(false)
        {
            EntryKind::Directory
        } else {
            EntryKind::File
        };
        let name = path
            .file_name()
            .map(|item| item.to_string_lossy().to_string())
            .unwrap_or_else(|| path.display().to_string());
        let metadata = if view_options.requires_metadata() {
            Some(metadata)
        } else {
            None
        };

        Ok(Self {
            name: name.clone(),
            path: Some(path),
            link_target,
            network_resource: None,
            kind,
            size_bytes: if view_options.show_size && matches!(kind, EntryKind::File) {
                metadata.as_ref().map(|item| item.file_size())
            } else {
                None
            },
            type_label: if shortcut {
                Some("skrót".to_string())
            } else if file_type.is_symlink() {
                if matches!(kind, EntryKind::Directory) {
                    Some("link do katalogu".to_string())
                } else {
                    Some("link symboliczny".to_string())
                }
            } else {
                match kind {
                    EntryKind::Directory => Some("katalog".to_string()),
                    EntryKind::File => Some(infer_file_type(&name)),
                    EntryKind::Drive => Some("dysk".to_string()),
                    EntryKind::NetworkPlaceholder => Some("zasób sieciowy".to_string()),
                    _ => None,
                }
            },
            created_label: if view_options.show_created {
                metadata
                    .as_ref()
                    .and_then(|item| format_filetime_label(item.creation_time()))
            } else {
                None
            },
            modified_label: if view_options.show_modified {
                metadata
                    .as_ref()
                    .and_then(|item| format_filetime_label(item.last_write_time()))
            } else {
                None
            },
        })
    }

    fn network_resource(resource: NetworkResource) -> Self {
        let name = resource.effective_display_name();
        Self {
            name,
            path: None,
            link_target: None,
            network_resource: Some(resource.clone()),
            kind: EntryKind::NetworkResource,
            size_bytes: None,
            type_label: Some(resource.protocol.inferred_type_label().to_string()),
            created_label: None,
            modified_label: None,
        }
    }

    fn stable_key(&self) -> String {
        match &self.path {
            Some(path) => format!("{:?}:{}", self.kind, path.display()),
            None => format!("{:?}:{}", self.kind, self.name),
        }
    }

    fn is_markable(&self) -> bool {
        matches!(self.kind, EntryKind::Directory | EntryKind::File)
    }

    fn is_operable(&self) -> bool {
        self.is_markable()
    }

    fn spoken_details(&self, view_options: ViewOptions) -> Vec<String> {
        let mut details = Vec::new();
        if view_options.show_size && matches!(self.kind, EntryKind::File) {
            if let Some(size) = self.size_bytes {
                details.push(format!("rozmiar {}", format_bytes(size)));
            }
        }
        if view_options.show_type && matches!(self.kind, EntryKind::File) {
            if let Some(kind) = &self.type_label {
                details.push(format!("typ {kind}"));
            }
        }
        if view_options.show_created {
            if let Some(created) = &self.created_label {
                details.push(format!("data utworzenia {created}"));
            }
        }
        if view_options.show_modified {
            if let Some(modified) = &self.modified_label {
                details.push(format!("data modyfikacji {modified}"));
            }
        }
        details
    }

    fn visual_details(&self, view_options: ViewOptions) -> Vec<String> {
        let mut details = Vec::new();
        if view_options.show_size {
            if let Some(size) = self.size_bytes {
                details.push(format!("rozmiar {}", format_bytes(size)));
            }
        }
        if view_options.show_type {
            if let Some(kind) = &self.type_label {
                details.push(format!("typ {kind}"));
            }
        }
        if view_options.show_created {
            if let Some(created) = &self.created_label {
                details.push(format!("utw. {created}"));
            }
        }
        if view_options.show_modified {
            if let Some(modified) = &self.modified_label {
                details.push(format!("modyf. {modified}"));
            }
        }
        details
    }

    fn accessible_label(&self, marked: bool, view_options: ViewOptions) -> String {
        let mut base = match self.kind {
            EntryKind::Loading => "ładowanie katalogu".to_string(),
            EntryKind::SearchLoading => "trwa wyszukiwanie".to_string(),
            EntryKind::SearchEmpty => "brak wyników".to_string(),
            EntryKind::GoToDrives => "przejdź do dysków".to_string(),
            EntryKind::NetworkResource => self
                .network_resource
                .as_ref()
                .map(NetworkResource::summary_line)
                .unwrap_or_else(|| self.name.clone()),
            EntryKind::FavoriteDirectoriesRoot => "ulubione katalogi".to_string(),
            EntryKind::FavoriteFilesRoot => "ulubione pliki".to_string(),
            EntryKind::Parent => "katalog nadrzędny".to_string(),
            EntryKind::Directory => format!("{} katalog", self.name),
            EntryKind::File
                if self.link_target.is_some() || is_windows_shortcut_name(&self.name) =>
            {
                format!("{} skrót", self.name)
            }
            EntryKind::File => self.name.clone(),
            EntryKind::Drive => format!("dysk {}", self.name),
            EntryKind::NetworkPlaceholder => "zasoby sieciowe".to_string(),
        };
        let details = self.spoken_details(view_options);
        if !details.is_empty() {
            base.push_str(", ");
            base.push_str(&details.join(", "));
        }

        if marked && self.is_markable() {
            format!("zaznaczone {base}")
        } else {
            base
        }
    }

    fn visual_label(&self, marked: bool, view_options: ViewOptions) -> String {
        let prefix = if marked { "[x]" } else { "[ ]" };
        let mut label = match self.kind {
            EntryKind::Loading => "[...] Ładowanie katalogu...".to_string(),
            EntryKind::SearchLoading => "[SZUK] Trwa wyszukiwanie".to_string(),
            EntryKind::SearchEmpty => "[BRAK] Brak wyników".to_string(),
            EntryKind::GoToDrives => "[DRV] Przejdź do dysków".to_string(),
            EntryKind::NetworkResource => format!(
                "[NET] {}",
                self.network_resource
                    .as_ref()
                    .map(NetworkResource::summary_line)
                    .unwrap_or_else(|| self.name.clone())
            ),
            EntryKind::FavoriteDirectoriesRoot => "[ULUB] Ulubione katalogi".to_string(),
            EntryKind::FavoriteFilesRoot => "[ULUB] Ulubione pliki".to_string(),
            EntryKind::Parent => "[..] Katalog nadrzędny".to_string(),
            EntryKind::Directory => format!("{prefix} {}\\", self.name),
            EntryKind::File
                if self.link_target.is_some() || is_windows_shortcut_name(&self.name) =>
            {
                format!("{prefix} {} ->", self.name)
            }
            EntryKind::File => format!("{prefix} {}", self.name),
            EntryKind::Drive => format!("[DYSK] {}", self.name),
            EntryKind::NetworkPlaceholder => "[NET] Zasoby sieciowe".to_string(),
        };
        let details = self.visual_details(view_options);
        if !details.is_empty() {
            label.push_str(" | ");
            label.push_str(&details.join(" | "));
        }
        label
    }
}

struct PanelModel {
    title: &'static str,
    location: PanelLocation,
    last_filesystem_path: PathBuf,
    entries: Vec<PanelEntry>,
    selected: usize,
    marked: HashSet<PathBuf>,
    label_hwnd: HWND,
    list_hwnd: HWND,
    loading: bool,
    load_generation: u64,
    search_generation: u64,
    pending_select: Option<PathBuf>,
    pending_key: Option<String>,
    search_state: Option<SearchState>,
    search_in_progress: bool,
}

impl PanelModel {
    fn new(title: &'static str, start_path: PathBuf) -> io::Result<Self> {
        let mut panel = Self {
            title,
            location: PanelLocation::Filesystem(start_path.clone()),
            last_filesystem_path: start_path,
            entries: Vec::new(),
            selected: 0,
            marked: HashSet::new(),
            label_hwnd: null_mut(),
            list_hwnd: null_mut(),
            loading: false,
            load_generation: 0,
            search_generation: 0,
            pending_select: None,
            pending_key: None,
            search_state: None,
            search_in_progress: false,
        };
        panel.entries = Self::loading_entries(&panel.location);
        Ok(panel)
    }

    fn base_entries(location: &PanelLocation) -> Vec<PanelEntry> {
        match location {
            PanelLocation::Drives => Vec::new(),
            PanelLocation::NetworkResources => vec![PanelEntry::go_to_drives()],
            PanelLocation::Filesystem(path) => {
                if let Some(parent) = path.parent() {
                    vec![PanelEntry::parent(parent.to_path_buf())]
                } else {
                    vec![PanelEntry::go_to_drives()]
                }
            }
            PanelLocation::Remote(remote) => {
                if let Some(parent) = remote_parent_path(&remote.path) {
                    vec![PanelEntry::parent(parent)]
                } else {
                    vec![PanelEntry::network_placeholder()]
                }
            }
            PanelLocation::Archive(archive) => {
                if archive.inside_path.as_os_str().is_empty() {
                    vec![PanelEntry::parent(archive.archive_path.clone())]
                } else {
                    vec![PanelEntry::parent(
                        archive
                            .inside_path
                            .parent()
                            .map(Path::to_path_buf)
                            .unwrap_or_default(),
                    )]
                }
            }
            PanelLocation::FavoriteDirectories | PanelLocation::FavoriteFiles => {
                vec![PanelEntry::go_to_drives()]
            }
        }
    }

    fn loading_entries(location: &PanelLocation) -> Vec<PanelEntry> {
        let mut entries = Self::base_entries(location);
        entries.push(PanelEntry::loading());
        entries
    }

    fn current_dir(&self) -> Option<&Path> {
        match &self.location {
            PanelLocation::Filesystem(path) => Some(path.as_path()),
            PanelLocation::Remote(_) | PanelLocation::Archive(_) => None,
            PanelLocation::Drives
            | PanelLocation::NetworkResources
            | PanelLocation::FavoriteDirectories
            | PanelLocation::FavoriteFiles => None,
        }
    }

    fn current_dir_owned(&self) -> Option<PathBuf> {
        self.current_dir().map(Path::to_path_buf)
    }

    fn current_remote_dir_owned(&self) -> Option<PathBuf> {
        match &self.location {
            PanelLocation::Remote(remote) => Some(remote.path.clone()),
            _ => None,
        }
    }

    fn location_label(&self) -> String {
        match &self.location {
            PanelLocation::Filesystem(path) => {
                if let Some(search) = &self.search_state {
                    format!(
                        "{}: {} [{}{}: {}]",
                        self.title,
                        path.display(),
                        if self.search_in_progress {
                            "trwa wyszukiwanie"
                        } else {
                            "wyniki wyszukiwania"
                        },
                        if search.recursive {
                            ", rekurencyjnie"
                        } else {
                            ""
                        },
                        search.pattern
                    )
                } else if self.loading {
                    format!("{}: {} [ładowanie]", self.title, path.display())
                } else {
                    format!("{}: {}", self.title, path.display())
                }
            }
            PanelLocation::Remote(remote) => {
                let suffix = remote_display_suffix(&remote.path);
                if self.loading {
                    format!(
                        "{}: {}{} [ładowanie]",
                        self.title,
                        remote.resource.effective_display_name(),
                        suffix
                    )
                } else {
                    format!(
                        "{}: {}{}",
                        self.title,
                        remote.resource.effective_display_name(),
                        suffix
                    )
                }
            }
            PanelLocation::Archive(archive) => {
                let inside = archive_internal_display_path(&archive.inside_path);
                if self.loading {
                    format!(
                        "{}: archiwum {}{} [ładowanie]",
                        self.title,
                        archive.archive_path.display(),
                        inside
                    )
                } else {
                    format!(
                        "{}: archiwum {}{}",
                        self.title,
                        archive.archive_path.display(),
                        inside
                    )
                }
            }
            PanelLocation::Drives => {
                if self.loading {
                    format!("{}: Dyski i zasoby sieciowe [ładowanie]", self.title)
                } else {
                    format!("{}: Dyski i zasoby sieciowe", self.title)
                }
            }
            PanelLocation::NetworkResources => {
                if self.loading {
                    format!("{}: Zasoby sieciowe [ładowanie]", self.title)
                } else {
                    format!("{}: Zasoby sieciowe", self.title)
                }
            }
            PanelLocation::FavoriteDirectories => {
                if self.loading {
                    format!("{}: Ulubione katalogi [ładowanie]", self.title)
                } else {
                    format!("{}: Ulubione katalogi", self.title)
                }
            }
            PanelLocation::FavoriteFiles => {
                if self.loading {
                    format!("{}: Ulubione pliki [ładowanie]", self.title)
                } else {
                    format!("{}: Ulubione pliki", self.title)
                }
            }
        }
    }

    fn selected_entry(&self) -> Option<&PanelEntry> {
        self.entries.get(self.selected)
    }

    fn is_search_active(&self) -> bool {
        self.search_state.is_some()
    }

    fn set_selection(&mut self, index: usize) {
        if self.entries.is_empty() {
            self.selected = 0;
        } else {
            self.selected = index.min(self.entries.len() - 1);
        }
    }

    fn select_path(&mut self, path: &Path) {
        if let Some(index) = self
            .entries
            .iter()
            .position(|entry| entry.path.as_deref() == Some(path))
        {
            self.selected = index;
        }
    }

    fn selection_announcement(&self, view_options: ViewOptions) -> String {
        let mut message = self
            .selected_entry()
            .map(|entry| {
                let marked = entry
                    .path
                    .as_ref()
                    .map(|path| self.marked.contains(path))
                    .unwrap_or(false);
                entry.accessible_label(marked, view_options)
            })
            .unwrap_or_else(|| "brak elementów".to_string());

        if !self.entries.is_empty() && self.selected == self.entries.len().saturating_sub(1) {
            message.push_str(", koniec listy");
        }

        message
    }
}

#[derive(Clone, Copy)]
enum PanelAction {
    Open,
    Up,
    ToggleMark,
    ClipboardCopy,
    ClipboardCut,
    ClipboardPaste,
    UnmarkAll,
    InvertMarks,
    SwitchPanel,
    ContextMenu,
    Rename,
    Copy,
    Move,
    NewFolder,
    Delete,
    Refresh,
    SelectFirst,
    SelectLast,
    MarkAll,
    Search,
    ExitSearch,
    MarkByExtension,
    MarkByName,
}

impl PanelAction {
    fn from_wparam(value: WPARAM) -> Option<Self> {
        match value as u32 {
            1 => Some(Self::Open),
            2 => Some(Self::Up),
            3 => Some(Self::ToggleMark),
            4 => Some(Self::ClipboardCopy),
            5 => Some(Self::ClipboardCut),
            6 => Some(Self::ClipboardPaste),
            7 => Some(Self::UnmarkAll),
            8 => Some(Self::InvertMarks),
            9 => Some(Self::SwitchPanel),
            10 => Some(Self::ContextMenu),
            11 => Some(Self::Rename),
            12 => Some(Self::Copy),
            13 => Some(Self::Move),
            14 => Some(Self::NewFolder),
            15 => Some(Self::Delete),
            16 => Some(Self::Refresh),
            17 => Some(Self::SelectFirst),
            18 => Some(Self::SelectLast),
            19 => Some(Self::MarkAll),
            20 => Some(Self::Search),
            21 => Some(Self::ExitSearch),
            22 => Some(Self::MarkByExtension),
            23 => Some(Self::MarkByName),
            _ => None,
        }
    }

    const fn as_wparam(self) -> WPARAM {
        match self {
            Self::Open => 1,
            Self::Up => 2,
            Self::ToggleMark => 3,
            Self::ClipboardCopy => 4,
            Self::ClipboardCut => 5,
            Self::ClipboardPaste => 6,
            Self::UnmarkAll => 7,
            Self::InvertMarks => 8,
            Self::SwitchPanel => 9,
            Self::ContextMenu => 10,
            Self::Rename => 11,
            Self::Copy => 12,
            Self::Move => 13,
            Self::NewFolder => 14,
            Self::Delete => 15,
            Self::Refresh => 16,
            Self::SelectFirst => 17,
            Self::SelectLast => 18,
            Self::MarkAll => 19,
            Self::Search => 20,
            Self::ExitSearch => 21,
            Self::MarkByExtension => 22,
            Self::MarkByName => 23,
        }
    }
}

#[derive(Default, Clone, Copy)]
struct ItemCounts {
    files: usize,
    directories: usize,
    bytes: u64,
}

impl ItemCounts {
    fn add(&mut self, other: Self) {
        self.files += other.files;
        self.directories += other.directories;
        self.bytes = self.bytes.saturating_add(other.bytes);
    }

    fn added(mut self, other: Self) -> Self {
        self.add(other);
        self
    }

    fn scaled(self, ratio: f64) -> Self {
        let ratio = ratio.clamp(0.0, 1.0);
        Self {
            files: ((self.files as f64) * ratio).round() as usize,
            directories: ((self.directories as f64) * ratio).round() as usize,
            bytes: ((self.bytes as f64) * ratio).round() as u64,
        }
    }
}

struct InputDialogState {
    owner: HWND,
    prompt: String,
    prompt_lines: Vec<String>,
    initial: String,
    result: Option<String>,
    done: bool,
    accepted: bool,
    prompt_hwnd: HWND,
    edit_hwnd: HWND,
    ok_hwnd: HWND,
    cancel_hwnd: HWND,
}

struct SearchDialogState {
    owner: HWND,
    result: Option<(String, bool)>,
    done: bool,
    accepted: bool,
    prompt_lines: Vec<String>,
    prompt_hwnd: HWND,
    edit_hwnd: HWND,
    local_hwnd: HWND,
    recursive_hwnd: HWND,
    cancel_hwnd: HWND,
}

struct NetworkDialogState {
    owner: HWND,
    initial: NetworkResource,
    result: Option<NetworkResource>,
    done: bool,
    accepted: bool,
    prompt_lines: Vec<String>,
    info_hwnd: HWND,
    protocol_hwnd: HWND,
    host_hwnd: HWND,
    username_hwnd: HWND,
    password_hwnd: HWND,
    ssh_key_hwnd: HWND,
    ssh_key_browse_hwnd: HWND,
    directory_hwnd: HWND,
    display_name_hwnd: HWND,
    anonymous_hwnd: HWND,
    ok_hwnd: HWND,
    cancel_hwnd: HWND,
}

struct DiscoveryDialogState {
    owner: HWND,
    result: Option<DiscoveredServer>,
    done: bool,
    servers: Vec<DiscoveredServer>,
    prompt_lines: Vec<String>,
    info_hwnd: HWND,
    list_hwnd: HWND,
    add_hwnd: HWND,
    cancel_hwnd: HWND,
}

#[derive(Clone)]
struct ArchiveCreateOptions {
    format: ArchiveFormat,
    name: String,
    compression_level: u8,
    encrypted: bool,
    encryption: String,
    password: String,
    volume_size: String,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum ArchiveFormat {
    SevenZip,
    Zip,
    Tar,
    TarGz,
    TarBz2,
    TarXz,
    Gzip,
    Bzip2,
    Xz,
    Wim,
}

impl ArchiveFormat {
    const ALL: [Self; 10] = [
        Self::SevenZip,
        Self::Zip,
        Self::Tar,
        Self::TarGz,
        Self::TarBz2,
        Self::TarXz,
        Self::Gzip,
        Self::Bzip2,
        Self::Xz,
        Self::Wim,
    ];

    const fn label(self) -> &'static str {
        match self {
            Self::SevenZip => "7z",
            Self::Zip => "zip",
            Self::Tar => "tar",
            Self::TarGz => "tar.gz",
            Self::TarBz2 => "tar.bz2",
            Self::TarXz => "tar.xz",
            Self::Gzip => "gzip (.gz, jeden plik)",
            Self::Bzip2 => "bzip2 (.bz2, jeden plik)",
            Self::Xz => "xz (.xz, jeden plik)",
            Self::Wim => "wim",
        }
    }

    const fn extension(self) -> &'static str {
        match self {
            Self::SevenZip => "7z",
            Self::Zip => "zip",
            Self::Tar => "tar",
            Self::TarGz => "tar.gz",
            Self::TarBz2 => "tar.bz2",
            Self::TarXz => "tar.xz",
            Self::Gzip => "gz",
            Self::Bzip2 => "bz2",
            Self::Xz => "xz",
            Self::Wim => "wim",
        }
    }

    const fn seven_zip_type(self) -> &'static str {
        match self {
            Self::SevenZip => "7z",
            Self::Zip => "zip",
            Self::Tar | Self::TarGz | Self::TarBz2 | Self::TarXz => "tar",
            Self::Gzip => "gzip",
            Self::Bzip2 => "bzip2",
            Self::Xz => "xz",
            Self::Wim => "wim",
        }
    }

    const fn compressed_tar_codec(self) -> Option<&'static str> {
        match self {
            Self::TarGz => Some("gzip"),
            Self::TarBz2 => Some("bzip2"),
            Self::TarXz => Some("xz"),
            _ => None,
        }
    }

    const fn is_single_file_compressor(self) -> bool {
        matches!(self, Self::Gzip | Self::Bzip2 | Self::Xz)
    }

    const fn supports_encryption(self) -> bool {
        matches!(self, Self::SevenZip | Self::Zip)
    }

    fn from_index(index: usize) -> Self {
        Self::ALL.get(index).copied().unwrap_or(Self::SevenZip)
    }
}

#[derive(Clone, Copy)]
enum ExtractMode {
    Here,
    NamedFolder,
    OtherPanel,
}

struct ArchiveCreateDialogState {
    owner: HWND,
    initial_name: String,
    result: Option<ArchiveCreateOptions>,
    done: bool,
    accepted: bool,
    prompt_lines: Vec<String>,
    info_hwnd: HWND,
    format_hwnd: HWND,
    name_hwnd: HWND,
    level_hwnd: HWND,
    encrypted_hwnd: HWND,
    encryption_hwnd: HWND,
    password_hwnd: HWND,
    volume_hwnd: HWND,
    ok_hwnd: HWND,
    cancel_hwnd: HWND,
}

struct OperationPromptState {
    owner: HWND,
    lines: Vec<String>,
    done: bool,
    result: i32,
    buttons: Vec<DialogButton>,
    info_hwnd: HWND,
    button_hwnds: Vec<HWND>,
}

struct ProgressDialogState {
    owner: HWND,
    lines: Vec<String>,
    done: bool,
    running: bool,
    auto_close_on_success: bool,
    auto_close_on_retryable_error: bool,
    cancel_flag: Arc<AtomicBool>,
    receiver: Receiver<ProgressEvent>,
    conflict_response_sender: Sender<ConflictChoice>,
    shared_outcome: Arc<Mutex<WorkerOutcome>>,
    info_hwnd: HWND,
    action_hwnd: HWND,
    current_line: usize,
}

#[derive(Clone, Copy)]
struct DialogButton {
    id: i32,
    label: &'static str,
    is_default: bool,
}

#[derive(Clone, Copy, Default)]
struct ProgressDialogOptions {
    auto_close_on_success: bool,
    auto_close_on_retryable_error: bool,
}

#[derive(Clone)]
struct EntryBatch {
    directories: Vec<PanelEntry>,
    files: Vec<PanelEntry>,
}

impl EntryBatch {
    fn new() -> Self {
        Self {
            directories: Vec::new(),
            files: Vec::new(),
        }
    }

    fn into_entries(self) -> Vec<PanelEntry> {
        let mut entries = self.directories;
        entries.extend(self.files);
        entries
    }
}

struct PanelLoadMessage {
    panel_index: usize,
    generation: u64,
    entries: Vec<PanelEntry>,
    first_chunk: bool,
    last_chunk: bool,
    error: Option<String>,
}

struct SearchMessage {
    panel_index: usize,
    generation: u64,
    entries: Vec<PanelEntry>,
    preferred_path: Option<PathBuf>,
    match_count: usize,
    error: Option<String>,
}

struct AppState {
    hwnd: HWND,
    panels: [PanelModel; 2],
    active_panel: usize,
    status_hwnd: HWND,
    status_line: String,
    nvda: NvdaController,
    font: HFONT,
    background_brush: HBRUSH,
    panel_load_sender: Sender<PanelLoadMessage>,
    panel_load_receiver: Receiver<PanelLoadMessage>,
    search_sender: Sender<SearchMessage>,
    search_receiver: Receiver<SearchMessage>,
    settings: AppSettings,
    view_options: ViewOptions,
    discovery_cache: DiscoveryCache,
    clipboard: Option<AppClipboardState>,
}

#[repr(C)]
struct DropFilesHeader {
    p_files: u32,
    pt_x: i32,
    pt_y: i32,
    f_nc: i32,
    f_wide: i32,
}

struct ProgressReporter {
    sender: Sender<ProgressEvent>,
    notify_hwnd: HWND,
    cancel_flag: Arc<AtomicBool>,
    action: &'static str,
    total: ItemCounts,
    processed: ItemCounts,
    started_at: Instant,
    last_sent: Instant,
    last_lines: Vec<String>,
    status: String,
    current: Option<PathBuf>,
    destination: Option<PathBuf>,
    conflict_response_receiver: Receiver<ConflictChoice>,
}

#[derive(Clone)]
enum ProgressEvent {
    Update(Vec<String>),
    Conflict {
        destination: PathBuf,
    },
    Finished {
        lines: Vec<String>,
        speech: String,
        outcome: WorkerOutcome,
    },
}

#[derive(Clone, PartialEq, Eq)]
enum WorkerOutcome {
    Success,
    Canceled,
    Error(String),
}

#[derive(Clone, Debug, Serialize, Deserialize)]
enum ElevatedLocalOperation {
    Rename {
        source: PathBuf,
        destination: PathBuf,
    },
    CreateDir {
        path: PathBuf,
    },
    DeleteTargets {
        targets: Vec<PathBuf>,
    },
    CopyMove {
        sources: Vec<PathBuf>,
        destination_dir: PathBuf,
        move_mode: bool,
    },
}

#[derive(Clone, Debug, Serialize, Deserialize)]
struct ElevatedLocalRequestFile {
    operation: ElevatedLocalOperation,
    success: Option<bool>,
    error: Option<String>,
}

impl AppState {
    fn new() -> io::Result<Self> {
        let settings = load_settings();
        let current_dir = std::env::current_dir().unwrap_or_else(|_| PathBuf::from("C:\\"));
        let left = PanelModel::new("LEWY", current_dir.clone())?;
        let right = PanelModel::new("PRAWY", current_dir)?;
        let nvda = NvdaController::new();
        let (panel_load_sender, panel_load_receiver) = mpsc::channel();
        let (search_sender, search_receiver) = mpsc::channel();
        let status_line = if nvda.is_available() {
            "NVDA podłączone".to_string()
        } else {
            "NVDA niedostępne".to_string()
        };

        let font = unsafe {
            CreateFontW(
                -20,
                0,
                0,
                0,
                FW_BOLD as i32,
                0,
                0,
                0,
                DEFAULT_CHARSET.into(),
                OUT_DEFAULT_PRECIS.into(),
                CLIP_DEFAULT_PRECIS.into(),
                CLEARTYPE_QUALITY.into(),
                (DEFAULT_PITCH | FF_MODERN) as u32,
                wide("Consolas").as_ptr(),
            )
        };

        let background_brush = unsafe { CreateSolidBrush(BLACK) };

        Ok(Self {
            hwnd: null_mut(),
            panels: [left, right],
            active_panel: 0,
            status_hwnd: null_mut(),
            status_line,
            nvda,
            font,
            background_brush,
            panel_load_sender,
            panel_load_receiver,
            search_sender,
            search_receiver,
            settings: settings.clone(),
            view_options: settings.view_options,
            discovery_cache: settings.discovery_cache.clone(),
            clipboard: None,
        })
    }

    fn save_settings(&mut self) -> io::Result<()> {
        self.settings.view_options = self.view_options;
        self.settings.discovery_cache = self.discovery_cache.clone();
        save_settings_file(&self.settings)
    }

    fn save_settings_or_report(&mut self) {
        if let Err(error) = self.save_settings() {
            self.report_error(error);
        }
    }

    unsafe fn initialize_window(&mut self, hwnd: HWND) -> Result<(), String> {
        self.hwnd = hwnd;
        self.create_menu();
        self.create_controls()?;
        self.refresh_all().map_err(|error| error.to_string())?;
        self.update_status_control();
        SetFocus(self.panels[0].list_hwnd);
        self.announce_selection();
        Ok(())
    }

    unsafe fn create_menu(&self) {
        let menu = CreateMenu();
        let file_menu = CreatePopupMenu();
        let edit_menu = CreatePopupMenu();
        let mark_menu = CreatePopupMenu();
        let view_menu = CreatePopupMenu();
        let options_menu = CreatePopupMenu();
        let old_menu = if self.hwnd.is_null() {
            null_mut()
        } else {
            GetMenu(self.hwnd)
        };
        AppendMenuW(
            file_menu,
            MF_STRING,
            IDM_RENAME as usize,
            wide("Zmień nazwę\tF2").as_ptr(),
        );
        AppendMenuW(
            file_menu,
            MF_STRING,
            IDM_NEW_FOLDER as usize,
            wide("Nowy katalog\tF7").as_ptr(),
        );
        AppendMenuW(
            file_menu,
            MF_STRING,
            IDM_DELETE as usize,
            wide("Usuń\tDelete").as_ptr(),
        );
        AppendMenuW(file_menu, MF_SEPARATOR, 0, null());
        AppendMenuW(
            file_menu,
            MF_STRING,
            IDM_PROPERTIES as usize,
            wide("Właściwości").as_ptr(),
        );
        AppendMenuW(
            file_menu,
            MF_STRING,
            IDM_PERMISSIONS as usize,
            wide("Uprawnienia").as_ptr(),
        );
        AppendMenuW(file_menu, MF_SEPARATOR, 0, null());
        AppendMenuW(
            file_menu,
            MF_STRING,
            IDM_EXTRACT_HERE as usize,
            wide("Wypakuj tutaj").as_ptr(),
        );
        AppendMenuW(
            file_menu,
            MF_STRING,
            IDM_EXTRACT_TO_FOLDER as usize,
            wide("Wypakuj do katalogu z nazwą archiwum").as_ptr(),
        );
        AppendMenuW(
            file_menu,
            MF_STRING,
            IDM_EXTRACT_TO_OTHER_PANEL as usize,
            wide("Wypakuj do drugiego panelu").as_ptr(),
        );
        AppendMenuW(
            file_menu,
            MF_STRING,
            IDM_CREATE_ARCHIVE as usize,
            wide("Utwórz archiwum lub obraz dysku").as_ptr(),
        );
        AppendMenuW(
            file_menu,
            MF_STRING,
            IDM_JOIN_SPLIT_ARCHIVE as usize,
            wide("Połącz podzielone archiwum").as_ptr(),
        );
        AppendMenuW(file_menu, MF_SEPARATOR, 0, null());
        AppendMenuW(
            file_menu,
            MF_STRING,
            IDM_CHECKSUM_CREATE as usize,
            wide("Utwórz sumę kontrolną SHA-256").as_ptr(),
        );
        AppendMenuW(
            file_menu,
            MF_STRING,
            IDM_CHECKSUM_VERIFY as usize,
            wide("Sprawdź sumę kontrolną SHA-256").as_ptr(),
        );
        AppendMenuW(file_menu, MF_SEPARATOR, 0, null());
        AppendMenuW(
            file_menu,
            MF_STRING,
            IDM_ADD_NETWORK_CONNECTION as usize,
            wide("Dodaj połączenie sieciowe").as_ptr(),
        );
        AppendMenuW(
            file_menu,
            MF_STRING,
            IDM_DISCOVER_NETWORK_SERVERS as usize,
            wide("Wyszukaj serwery w sieci lokalnej").as_ptr(),
        );
        AppendMenuW(
            file_menu,
            MF_STRING,
            IDM_EDIT_NETWORK_CONNECTION as usize,
            wide("Edytuj połączenie sieciowe").as_ptr(),
        );
        AppendMenuW(
            file_menu,
            MF_STRING,
            IDM_REMOVE_NETWORK_CONNECTION as usize,
            wide("Usuń połączenie sieciowe").as_ptr(),
        );
        AppendMenuW(file_menu, MF_SEPARATOR, 0, null());
        AppendMenuW(
            file_menu,
            MF_STRING,
            IDM_QUIT as usize,
            wide("Zamknij\tAlt+F4").as_ptr(),
        );

        AppendMenuW(
            edit_menu,
            MF_STRING,
            IDM_COPY as usize,
            wide("Kopiuj\tF5").as_ptr(),
        );
        AppendMenuW(
            edit_menu,
            MF_STRING,
            IDM_MOVE as usize,
            wide("Przenieś\tF6").as_ptr(),
        );
        AppendMenuW(
            edit_menu,
            MF_STRING,
            IDM_SEARCH as usize,
            wide("Wyszukaj\tCtrl+F").as_ptr(),
        );

        AppendMenuW(
            mark_menu,
            MF_STRING,
            IDM_MARK_ALL as usize,
            wide("Zaznacz wszystko\tCtrl+A").as_ptr(),
        );
        AppendMenuW(
            mark_menu,
            MF_STRING,
            IDM_UNMARK_ALL as usize,
            wide("Odznacz wszystko").as_ptr(),
        );
        AppendMenuW(
            mark_menu,
            MF_STRING,
            IDM_INVERT_MARKS as usize,
            wide("Odwróć zaznaczenie").as_ptr(),
        );
        AppendMenuW(mark_menu, MF_SEPARATOR, 0, null());
        AppendMenuW(
            mark_menu,
            MF_STRING,
            IDM_MARK_EXTENSION as usize,
            wide("Zaznacz pliki po rozszerzeniu").as_ptr(),
        );
        AppendMenuW(
            mark_menu,
            MF_STRING,
            IDM_MARK_NAME as usize,
            wide("Zaznacz po nazwie lub regexie").as_ptr(),
        );
        AppendMenuW(
            view_menu,
            MF_STRING,
            IDM_REFRESH as usize,
            wide("Odśwież\tF12").as_ptr(),
        );
        AppendMenuW(
            view_menu,
            MF_STRING
                | if self.view_options.show_size {
                    MF_CHECKED
                } else {
                    MF_UNCHECKED
                },
            IDM_VIEW_SIZE as usize,
            wide("Pokaż rozmiar elementu").as_ptr(),
        );
        AppendMenuW(
            view_menu,
            MF_STRING
                | if self.view_options.show_type {
                    MF_CHECKED
                } else {
                    MF_UNCHECKED
                },
            IDM_VIEW_TYPE as usize,
            wide("Pokaż typ elementu").as_ptr(),
        );
        AppendMenuW(
            view_menu,
            MF_STRING
                | if self.view_options.show_created {
                    MF_CHECKED
                } else {
                    MF_UNCHECKED
                },
            IDM_VIEW_CREATED as usize,
            wide("Pokaż datę utworzenia").as_ptr(),
        );
        AppendMenuW(
            view_menu,
            MF_STRING
                | if self.view_options.show_modified {
                    MF_CHECKED
                } else {
                    MF_UNCHECKED
                },
            IDM_VIEW_MODIFIED as usize,
            wide("Pokaż datę modyfikacji").as_ptr(),
        );
        AppendMenuW(
            options_menu,
            MF_STRING,
            IDM_OPTIONS as usize,
            wide("Ustawienia").as_ptr(),
        );

        AppendMenuW(menu, MF_POPUP, file_menu as usize, wide("&Plik").as_ptr());
        AppendMenuW(menu, MF_POPUP, edit_menu as usize, wide("&Edycja").as_ptr());
        AppendMenuW(
            menu,
            MF_POPUP,
            mark_menu as usize,
            wide("&Zaznaczanie").as_ptr(),
        );
        AppendMenuW(menu, MF_POPUP, view_menu as usize, wide("&Widok").as_ptr());
        AppendMenuW(
            menu,
            MF_POPUP,
            options_menu as usize,
            wide("&Opcje").as_ptr(),
        );
        AppendMenuW(
            menu,
            MF_STRING,
            IDM_ABOUT as usize,
            wide("&O programie").as_ptr(),
        );
        SetMenu(self.hwnd, menu);
        DrawMenuBar(self.hwnd);
        if !old_menu.is_null() {
            DestroyMenu(old_menu);
        }
    }

    unsafe fn show_context_menu(&mut self, panel_index: usize) {
        self.activate_panel(panel_index, false);
        let menu = create_context_popup_menu();
        let mut rect: RECT = std::mem::zeroed();
        GetWindowRect(self.panels[panel_index].list_hwnd, &mut rect);
        let x = rect.left + 24;
        let y = rect.top + 24;
        let command = TrackPopupMenu(
            menu,
            TPM_LEFTALIGN | TPM_TOPALIGN | TPM_RETURNCMD,
            x,
            y,
            0,
            self.hwnd,
            null(),
        ) as u16;
        DestroyMenu(menu);
        if command != 0 {
            self.handle_command(command as WPARAM, 0);
        }
    }

    unsafe fn create_controls(&mut self) -> Result<(), String> {
        self.panels[0].label_hwnd = create_static(self.hwnd, IDC_LEFT_LABEL)?;
        self.panels[0].list_hwnd = create_listbox(self.hwnd, IDC_LEFT_LIST)?;
        self.panels[1].label_hwnd = create_static(self.hwnd, IDC_RIGHT_LABEL)?;
        self.panels[1].list_hwnd = create_listbox(self.hwnd, IDC_RIGHT_LIST)?;
        self.status_hwnd = create_static(self.hwnd, IDC_STATUS)?;

        for panel_index in 0..2 {
            SendMessageW(
                self.panels[panel_index].label_hwnd,
                0x0030,
                self.font as WPARAM,
                1,
            );
            SendMessageW(
                self.panels[panel_index].list_hwnd,
                0x0030,
                self.font as WPARAM,
                1,
            );
            SetWindowSubclass(
                self.panels[panel_index].list_hwnd,
                Some(listbox_subclass_proc),
                panel_index + 1,
                panel_index,
            );
        }

        SendMessageW(self.status_hwnd, 0x0030, self.font as WPARAM, 1);
        self.layout();
        Ok(())
    }

    unsafe fn layout(&self) {
        let mut rect = std::mem::zeroed();
        GetClientRect(self.hwnd, &mut rect);

        let width = rect.right - rect.left;
        let height = rect.bottom - rect.top;
        let gap = 8;
        let label_height = 28;
        let status_height = 28;
        let top = 10;
        let left = 10;
        let panel_width = (width - left * 2 - gap) / 2;
        let list_top = top + label_height + 6;
        let list_height = height - list_top - status_height - 16;

        MoveWindow(
            self.panels[0].label_hwnd,
            left,
            top,
            panel_width,
            label_height,
            1,
        );
        MoveWindow(
            self.panels[1].label_hwnd,
            left + panel_width + gap,
            top,
            panel_width,
            label_height,
            1,
        );
        MoveWindow(
            self.panels[0].list_hwnd,
            left,
            list_top,
            panel_width,
            list_height,
            1,
        );
        MoveWindow(
            self.panels[1].list_hwnd,
            left + panel_width + gap,
            list_top,
            panel_width,
            list_height,
            1,
        );
        MoveWindow(
            self.status_hwnd,
            left,
            height - status_height - 8,
            width - left * 2,
            status_height,
            1,
        );
    }

    fn update_status_control(&self) {
        if !self.status_hwnd.is_null() {
            unsafe {
                SetWindowTextW(self.status_hwnd, wide(&self.status_line).as_ptr());
            }
        }
    }

    fn set_status(&mut self, message: impl Into<String>) {
        self.status_line = message.into();
        self.update_status_control();
    }

    fn notify(&mut self, message: impl Into<String>) {
        let message = message.into();
        self.set_status(message.clone());
        self.nvda.speak(&message);
    }

    fn notify_non_interrupting(&mut self, message: impl Into<String>) {
        let message = message.into();
        self.set_status(message.clone());
        self.nvda.speak_non_interrupting(&message);
    }

    fn report_error<E: std::fmt::Display>(&mut self, error: E) {
        self.notify(format!("błąd: {error}"));
        unsafe {
            MessageBoxW(
                self.hwnd,
                wide(&error.to_string()).as_ptr(),
                wide("Błąd").as_ptr(),
                MB_OK | MB_ICONERROR,
            );
        }
    }

    fn announce_panel_for(&mut self, panel_index: usize) {
        self.notify(if panel_index == 0 {
            "lewy panel"
        } else {
            "prawy panel"
        });
    }

    fn announce_selection(&mut self) {
        let message = self.panels[self.active_panel].selection_announcement(self.view_options);
        self.notify(message);
    }

    unsafe fn start_panel_load(&mut self, panel_index: usize, preferred_path: Option<PathBuf>) {
        let (location, generation) = {
            let panel = &mut self.panels[panel_index];
            panel.load_generation = panel.load_generation.wrapping_add(1);
            panel.search_generation = panel.search_generation.wrapping_add(1);
            panel.search_in_progress = false;
            panel.loading = true;
            panel.pending_select = preferred_path;
            panel.pending_key = panel.selected_entry().map(PanelEntry::stable_key);
            panel.entries = PanelModel::loading_entries(&panel.location);
            panel.selected = 0;
            (panel.location.clone(), panel.load_generation)
        };
        self.rebuild_panel(panel_index);

        let sender = self.panel_load_sender.clone();
        let view_options = self.view_options;
        let hwnd = self.hwnd as usize;
        let favorite_directories = self.settings.favorite_directories.clone();
        let favorite_files = self.settings.favorite_files.clone();
        let network_resources = self.settings.network_resources.clone();

        thread::spawn(move || match location {
            PanelLocation::Drives => {
                let entries = read_drive_entries();
                let _ = sender.send(PanelLoadMessage {
                    panel_index,
                    generation,
                    entries,
                    first_chunk: true,
                    last_chunk: true,
                    error: None,
                });
                unsafe {
                    SendMessageW(hwnd as HWND, WM_PANEL_LOAD_EVENT, 0, 0);
                }
            }
            PanelLocation::NetworkResources => {
                let entries = read_network_resource_entries(&network_resources);
                let _ = sender.send(PanelLoadMessage {
                    panel_index,
                    generation,
                    entries,
                    first_chunk: true,
                    last_chunk: true,
                    error: None,
                });
                unsafe {
                    SendMessageW(hwnd as HWND, WM_PANEL_LOAD_EVENT, 0, 0);
                }
            }
            PanelLocation::FavoriteDirectories => {
                let entries = read_favorite_entries(&favorite_directories, true, view_options);
                let _ = sender.send(PanelLoadMessage {
                    panel_index,
                    generation,
                    entries,
                    first_chunk: true,
                    last_chunk: true,
                    error: None,
                });
                unsafe {
                    SendMessageW(hwnd as HWND, WM_PANEL_LOAD_EVENT, 0, 0);
                }
            }
            PanelLocation::FavoriteFiles => {
                let entries = read_favorite_entries(&favorite_files, false, view_options);
                let _ = sender.send(PanelLoadMessage {
                    panel_index,
                    generation,
                    entries,
                    first_chunk: true,
                    last_chunk: true,
                    error: None,
                });
                unsafe {
                    SendMessageW(hwnd as HWND, WM_PANEL_LOAD_EVENT, 0, 0);
                }
            }
            PanelLocation::Archive(archive) => {
                match read_archive_entries(&archive, view_options) {
                    Ok(entries) => {
                        let _ = sender.send(PanelLoadMessage {
                            panel_index,
                            generation,
                            entries,
                            first_chunk: true,
                            last_chunk: true,
                            error: None,
                        });
                    }
                    Err(error) => {
                        let _ = sender.send(PanelLoadMessage {
                            panel_index,
                            generation,
                            entries: Vec::new(),
                            first_chunk: true,
                            last_chunk: true,
                            error: Some(error.to_string()),
                        });
                    }
                }
                unsafe {
                    SendMessageW(hwnd as HWND, WM_PANEL_LOAD_EVENT, 0, 0);
                }
            }
            PanelLocation::Filesystem(path) => {
                let read_dir = match fs::read_dir(&path) {
                    Ok(read_dir) => read_dir,
                    Err(error) => {
                        let _ = sender.send(PanelLoadMessage {
                            panel_index,
                            generation,
                            entries: Vec::new(),
                            first_chunk: true,
                            last_chunk: true,
                            error: Some(error.to_string()),
                        });
                        unsafe {
                            SendMessageW(hwnd as HWND, WM_PANEL_LOAD_EVENT, 0, 0);
                        }
                        return;
                    }
                };

                let mut batch = EntryBatch::new();
                let mut file_entries = Vec::new();
                let mut first_chunk = true;

                let send_batch = |batch: EntryBatch, first_chunk: bool, last_chunk: bool| -> bool {
                    let entries = batch.into_entries();
                    if entries.is_empty() && !last_chunk {
                        return first_chunk;
                    }
                    let _ = sender.send(PanelLoadMessage {
                        panel_index,
                        generation,
                        entries,
                        first_chunk,
                        last_chunk,
                        error: None,
                    });
                    unsafe {
                        SendMessageW(hwnd as HWND, WM_PANEL_LOAD_EVENT, 0, 0);
                    }
                    false
                };
                let mut already_completed = false;

                for entry in read_dir {
                    let Ok(entry) = entry else {
                        continue;
                    };
                    let Ok(panel_entry) = PanelEntry::from_dir_entry(entry, view_options) else {
                        continue;
                    };
                    match panel_entry.kind {
                        EntryKind::Directory => {
                            batch.directories.push(panel_entry);
                            if batch.directories.len() >= PANEL_LOAD_CHUNK_SIZE {
                                first_chunk = send_batch(
                                    std::mem::replace(&mut batch, EntryBatch::new()),
                                    first_chunk,
                                    false,
                                );
                            }
                        }
                        EntryKind::File => {
                            file_entries.push(panel_entry);
                        }
                        _ => {}
                    }
                }

                if !batch.directories.is_empty() {
                    first_chunk = send_batch(
                        std::mem::replace(&mut batch, EntryBatch::new()),
                        first_chunk,
                        file_entries.is_empty(),
                    );
                    already_completed = file_entries.is_empty();
                }

                for file_chunk in file_entries.chunks(PANEL_LOAD_CHUNK_SIZE) {
                    let _ = sender.send(PanelLoadMessage {
                        panel_index,
                        generation,
                        entries: file_chunk.to_vec(),
                        first_chunk,
                        last_chunk: false,
                        error: None,
                    });
                    first_chunk = false;
                    unsafe {
                        SendMessageW(hwnd as HWND, WM_PANEL_LOAD_EVENT, 0, 0);
                    }
                }

                if !already_completed {
                    let _ = sender.send(PanelLoadMessage {
                        panel_index,
                        generation,
                        entries: Vec::new(),
                        first_chunk,
                        last_chunk: true,
                        error: None,
                    });
                    unsafe {
                        SendMessageW(hwnd as HWND, WM_PANEL_LOAD_EVENT, 0, 0);
                    }
                }
            }
            PanelLocation::Remote(remote) => {
                match read_remote_entries(&remote, view_options) {
                    Ok(entries) => {
                        let _ = sender.send(PanelLoadMessage {
                            panel_index,
                            generation,
                            entries,
                            first_chunk: true,
                            last_chunk: true,
                            error: None,
                        });
                    }
                    Err(error) => {
                        let _ = sender.send(PanelLoadMessage {
                            panel_index,
                            generation,
                            entries: Vec::new(),
                            first_chunk: true,
                            last_chunk: true,
                            error: Some(error.to_string()),
                        });
                    }
                }
                unsafe {
                    SendMessageW(hwnd as HWND, WM_PANEL_LOAD_EVENT, 0, 0);
                }
            }
        });
    }

    unsafe fn process_panel_load_messages(&mut self) {
        while let Ok(message) = self.panel_load_receiver.try_recv() {
            if self.panels[message.panel_index].load_generation != message.generation {
                continue;
            }

            if let Some(error) = message.error {
                let location = self.panels[message.panel_index].location.clone();
                let retry_remote = match location {
                    PanelLocation::Remote(remote) => self.next_sftp_retry_remote(&remote, &error),
                    PanelLocation::Archive(_) => None,
                    _ => None,
                };
                if let Some(remote) = retry_remote {
                    self.panels[message.panel_index].loading = false;
                    self.panels[message.panel_index].location = PanelLocation::Remote(remote);
                    let _ = self.refresh_and_keep(message.panel_index, None);
                    continue;
                }
                let panel = &mut self.panels[message.panel_index];
                panel.loading = false;
                panel.entries = PanelModel::base_entries(&panel.location);
                self.rebuild_panel(message.panel_index);
                self.report_error(error);
                continue;
            }

            let panel = &mut self.panels[message.panel_index];
            let appended = message.entries;
            let mut requires_full_rebuild = false;
            let mut should_append = Vec::new();
            if message.first_chunk {
                panel.entries = PanelModel::base_entries(&panel.location);
                if !appended.is_empty() {
                    panel.entries.extend(appended.iter().cloned());
                }
                requires_full_rebuild = true;
            } else if !appended.is_empty() {
                panel.entries.extend(appended.iter().cloned());
                should_append = appended;
            }

            if message.last_chunk {
                panel.loading = false;
                panel.marked.retain(|path| {
                    panel
                        .entries
                        .iter()
                        .any(|entry| entry.path.as_deref() == Some(path.as_path()))
                });

                if let Some(path) = panel.pending_select.take() {
                    panel.select_path(&path);
                } else if let Some(key) = panel.pending_key.take() {
                    if let Some(index) = panel
                        .entries
                        .iter()
                        .position(|entry| entry.stable_key() == key)
                    {
                        panel.selected = index;
                    } else {
                        panel.selected = panel.selected.min(panel.entries.len().saturating_sub(1));
                    }
                } else {
                    panel.selected = panel.selected.min(panel.entries.len().saturating_sub(1));
                }

                if panel.entries.is_empty() {
                    panel.selected = 0;
                }
                requires_full_rebuild = requires_full_rebuild || message.first_chunk;
                let selected = panel.selected;
                let label = panel.location_label();
                let list_hwnd = panel.list_hwnd;
                let label_hwnd = panel.label_hwnd;
                if requires_full_rebuild {
                    self.rebuild_panel(message.panel_index);
                } else if !should_append.is_empty() {
                    self.append_entries_to_panel(message.panel_index, &should_append);
                    self.update_listbox_view(list_hwnd, selected);
                    SetWindowTextW(label_hwnd, wide(&label).as_ptr());
                } else {
                    self.update_listbox_view(list_hwnd, selected);
                    SetWindowTextW(label_hwnd, wide(&label).as_ptr());
                }
                if self.active_panel == message.panel_index {
                    self.focus_panel(message.panel_index, false);
                    self.announce_selection();
                }
            } else {
                if requires_full_rebuild {
                    self.rebuild_panel(message.panel_index);
                } else if !should_append.is_empty() {
                    self.append_entries_to_panel(message.panel_index, &should_append);
                }
            }
        }
    }

    unsafe fn process_search_messages(&mut self) {
        while let Ok(message) = self.search_receiver.try_recv() {
            let panel = &mut self.panels[message.panel_index];
            if panel.search_generation != message.generation {
                continue;
            }

            panel.search_in_progress = false;

            if let Some(error) = message.error {
                panel.entries = vec![PanelEntry::search_empty()];
                panel.selected = 0;
                self.rebuild_panel(message.panel_index);
                self.report_error(error);
                continue;
            }

            let mut entries = message.entries;
            if entries.is_empty() {
                entries.push(PanelEntry::search_empty());
            }

            panel.entries = entries;
            panel.marked.retain(|path| {
                panel
                    .entries
                    .iter()
                    .any(|entry| entry.path.as_deref() == Some(path.as_path()))
            });
            panel.selected = 0;
            if let Some(path) = message.preferred_path.as_deref() {
                panel.select_path(path);
            }

            self.rebuild_panel(message.panel_index);
            if self.active_panel == message.panel_index {
                self.focus_panel(message.panel_index, false);
                self.notify(format!(
                    "znaleziono {}",
                    pluralized_elements(message.match_count)
                ));
            }
        }
    }

    unsafe fn refresh_all(&mut self) -> io::Result<()> {
        self.refresh_panel_after_mutation(0, None)?;
        self.refresh_panel_after_mutation(1, None)?;
        Ok(())
    }

    unsafe fn rebuild_panel(&mut self, panel_index: usize) {
        let panel = &self.panels[panel_index];
        SetWindowTextW(panel.label_hwnd, wide(&panel.location_label()).as_ptr());
        SendMessageW(panel.list_hwnd, LB_RESETCONTENT, 0, 0);
        for entry in &panel.entries {
            let marked = entry
                .path
                .as_ref()
                .map(|path| panel.marked.contains(path))
                .unwrap_or(false);
            let text = wide(&entry.visual_label(marked, self.view_options));
            SendMessageW(panel.list_hwnd, LB_ADDSTRING, 0, text.as_ptr() as LPARAM);
        }
        self.update_listbox_view(panel.list_hwnd, panel.selected);
        InvalidateRect(panel.list_hwnd, null(), 1);
    }

    unsafe fn append_entries_to_panel(&self, panel_index: usize, entries: &[PanelEntry]) {
        let panel = &self.panels[panel_index];
        for entry in entries {
            let marked = entry
                .path
                .as_ref()
                .map(|path| panel.marked.contains(path))
                .unwrap_or(false);
            let text = wide(&entry.visual_label(marked, self.view_options));
            SendMessageW(panel.list_hwnd, LB_ADDSTRING, 0, text.as_ptr() as LPARAM);
        }
        InvalidateRect(panel.list_hwnd, null(), 1);
    }

    unsafe fn sync_selection_from_control(&mut self, panel_index: usize) {
        let current = SendMessageW(self.panels[panel_index].list_hwnd, LB_GETCURSEL, 0, 0);
        if current >= 0 {
            self.panels[panel_index].set_selection(current as usize);
        }
    }

    unsafe fn update_listbox_view(&self, list_hwnd: HWND, selected: usize) {
        SendMessageW(list_hwnd, LB_SETCURSEL, selected, 0);
        let top_index = selected.saturating_sub(2);
        SendMessageW(list_hwnd, LB_SETTOPINDEX, top_index, 0);
    }

    unsafe fn activate_panel(&mut self, panel_index: usize, announce: bool) {
        if self.active_panel != panel_index {
            self.active_panel = panel_index;
            if announce {
                self.announce_panel_for(panel_index);
            }
        }
    }

    unsafe fn focus_panel(&mut self, panel_index: usize, announce: bool) {
        self.activate_panel(panel_index, announce);
        SetFocus(self.panels[panel_index].list_hwnd);
    }

    unsafe fn handle_panel_action(&mut self, panel_index: usize, action: PanelAction) {
        self.activate_panel(panel_index, matches!(action, PanelAction::SwitchPanel));
        self.sync_selection_from_control(panel_index);
        match action {
            PanelAction::Open => self.open_selected(panel_index),
            PanelAction::Up => self.go_up(panel_index),
            PanelAction::ToggleMark => self.toggle_mark(panel_index),
            PanelAction::ClipboardCopy => self.copy_selection_to_clipboard(panel_index, false),
            PanelAction::ClipboardCut => self.copy_selection_to_clipboard(panel_index, true),
            PanelAction::ClipboardPaste => self.paste_clipboard_into_panel(panel_index),
            PanelAction::UnmarkAll => self.unmark_all(panel_index),
            PanelAction::InvertMarks => self.invert_marks(panel_index),
            PanelAction::SwitchPanel => self.focus_panel(1 - panel_index, true),
            PanelAction::ContextMenu => self.show_context_menu(panel_index),
            PanelAction::Rename => self.rename_selected(panel_index),
            PanelAction::Copy => self.copy_or_move_selected(panel_index, false),
            PanelAction::Move => self.copy_or_move_selected(panel_index, true),
            PanelAction::NewFolder => self.create_folder(panel_index),
            PanelAction::Delete => self.delete_selected(panel_index),
            PanelAction::Refresh => {
                if let Err(error) = self.refresh_and_keep(panel_index, None) {
                    self.report_error(error);
                }
            }
            PanelAction::SelectFirst => self.select_index(panel_index, 0, true),
            PanelAction::SelectLast => {
                let last = self.panels[panel_index].entries.len().saturating_sub(1);
                self.select_index(panel_index, last, true);
            }
            PanelAction::MarkAll => self.mark_all(panel_index),
            PanelAction::Search => self.show_search_dialog_for_panel(panel_index),
            PanelAction::ExitSearch => self.exit_search_results(panel_index),
            PanelAction::MarkByExtension => self.mark_by_extension(panel_index),
            PanelAction::MarkByName => self.mark_by_name_pattern(panel_index),
        }
    }

    unsafe fn select_index(&mut self, panel_index: usize, index: usize, announce: bool) {
        self.panels[panel_index].set_selection(index);
        self.update_listbox_view(
            self.panels[panel_index].list_hwnd,
            self.panels[panel_index].selected,
        );
        if self.active_panel != panel_index {
            self.active_panel = panel_index;
        }
        if announce {
            self.announce_selection();
        }
    }

    unsafe fn search_next(&mut self, panel_index: usize, ch: char) {
        let needle = ch.to_lowercase().to_string();
        let panel = &self.panels[panel_index];
        if panel.entries.is_empty() {
            return;
        }

        let start = panel.selected + 1;
        let len = panel.entries.len();
        let found = (0..len).find_map(|offset| {
            let index = (start + offset) % len;
            let entry = &panel.entries[index];
            if entry.name.to_lowercase().starts_with(&needle) {
                Some(index)
            } else {
                None
            }
        });

        if let Some(index) = found {
            self.select_index(panel_index, index, true);
        }
    }

    unsafe fn enter_directory(&mut self, panel_index: usize, path: PathBuf) {
        self.panels[panel_index].search_state = None;
        match &self.panels[panel_index].location {
            PanelLocation::Remote(current) => {
                self.panels[panel_index].location = PanelLocation::Remote(RemoteLocation {
                    resource: current.resource.clone(),
                    original_resource: current.original_resource.clone(),
                    path: path.clone(),
                    sftp_privilege_mode: current.sftp_privilege_mode,
                });
            }
            PanelLocation::Archive(current) => {
                self.panels[panel_index].location = PanelLocation::Archive(ArchiveLocation {
                    archive_path: current.archive_path.clone(),
                    inside_path: normalize_archive_inner_path(&path),
                });
            }
            _ => {
                self.panels[panel_index].location = PanelLocation::Filesystem(path.clone());
                self.panels[panel_index].last_filesystem_path = path.clone();
            }
        }
        if let Err(error) = self.refresh_and_keep(panel_index, None) {
            self.report_error(error);
        }
    }

    unsafe fn open_local_file_or_link(&mut self, panel_index: usize, entry: PanelEntry) {
        let Some(path) = entry.path else {
            return;
        };

        if is_windows_shortcut_path(&path) {
            match resolve_windows_shortcut(&path) {
                Ok(target) => {
                    if local_path_is_directory(&target) {
                        self.enter_directory(panel_index, target.clone());
                        self.notify(format!("otwieram skrót do {}", display_name(&target)));
                    } else {
                        match open_with_system(&target) {
                            Ok(()) => {
                                self.notify(format!("otwieram skrót do {}", display_name(&target)))
                            }
                            Err(error) => self.report_error(error),
                        }
                    }
                    return;
                }
                Err(error) => {
                    self.notify_non_interrupting(format!(
                        "nie udało się odczytać skrótu, otwieram normalnie: {error}"
                    ));
                }
            }
        }

        if let Some(target) = entry.link_target {
            if local_path_is_directory(&target) {
                self.enter_directory(panel_index, target.clone());
                self.notify(format!("otwieram link do {}", display_name(&target)));
                return;
            }
            if target.is_file() {
                match open_with_system(&target) {
                    Ok(()) => self.notify(format!("otwieram link do {}", display_name(&target))),
                    Err(error) => self.report_error(error),
                }
                return;
            }
        }

        if is_archive_file_path(&path) {
            self.panels[panel_index].search_state = None;
            self.panels[panel_index].location =
                PanelLocation::Archive(ArchiveLocation::root(path.clone()));
            if let Err(error) = self.refresh_and_keep(panel_index, None) {
                self.report_error(error);
            } else {
                self.notify(format!("otwieram archiwum {}", display_name(&path)));
            }
            return;
        }

        match open_with_system(&path) {
            Ok(()) => self.notify(format!("otwieram {}", display_name(&path))),
            Err(error) => self.report_error(error),
        }
    }

    unsafe fn open_remote_file_or_link(
        &mut self,
        panel_index: usize,
        remote: RemoteLocation,
        entry: PanelEntry,
    ) {
        let Some(path) = entry.path.clone() else {
            return;
        };

        if let Some(target) = entry.link_target.clone() {
            match resolve_remote_link_target(&remote, &target) {
                Ok((resolved, EntryKind::Directory)) => {
                    self.enter_directory(panel_index, resolved.clone());
                    self.notify(format!("otwieram link do {}", remote_shell_path(&resolved)));
                    return;
                }
                Ok((resolved, _)) => {
                    self.open_remote_file_in_background(remote, resolved, display_name(&target));
                    return;
                }
                Err(error) if is_sftp_retry_error(&error) => {
                    self.open_remote_file_in_background(remote, target, entry.name);
                    return;
                }
                Err(error) => {
                    self.report_error(error);
                    return;
                }
            }
        }

        match remote_path_is_directory(&remote, &path) {
            Ok(true) => {
                self.enter_directory(panel_index, path.clone());
                self.notify(format!("otwieram katalog {}", remote_shell_path(&path)));
                return;
            }
            Ok(false) => {}
            Err(error) if is_sftp_retry_error(&error) => {
                if let Some(candidate) = self.next_sftp_retry_remote(&remote, &error.to_string()) {
                    self.open_remote_file_or_link(panel_index, candidate, entry);
                    return;
                }
            }
            Err(_) => {}
        }

        self.open_remote_file_in_background(remote, path, entry.name);
    }

    unsafe fn open_selected(&mut self, panel_index: usize) {
        let entry = match self.panels[panel_index].selected_entry().cloned() {
            Some(entry) => entry,
            None => {
                self.notify("brak elementów");
                return;
            }
        };

        match entry.kind {
            EntryKind::Loading => {
                self.notify("ładowanie katalogu");
            }
            EntryKind::SearchLoading => {
                self.notify("trwa wyszukiwanie");
            }
            EntryKind::SearchEmpty => {
                self.notify("brak wyników");
            }
            EntryKind::GoToDrives => {
                self.panels[panel_index].location = PanelLocation::Drives;
                if let Err(error) = self.refresh_and_keep(panel_index, None) {
                    self.report_error(error);
                }
            }
            EntryKind::NetworkResource => {
                if let Some(resource) = entry.network_resource {
                    if let Some(location) = resource.as_remote_location() {
                        self.panels[panel_index].search_state = None;
                        self.panels[panel_index].location = PanelLocation::Remote(location);
                        if let Err(error) = self.refresh_and_keep(panel_index, None) {
                            self.report_error(error);
                        } else {
                            self.notify(format!(
                                "otwieram zasób {}",
                                resource.effective_display_name()
                            ));
                        }
                    } else if let Some(path) = resource.smb_filesystem_path() {
                        self.panels[panel_index].search_state = None;
                        self.panels[panel_index].location = PanelLocation::Filesystem(path.clone());
                        self.panels[panel_index].last_filesystem_path = path.clone();
                        if let Err(error) = self.refresh_and_keep(panel_index, None) {
                            self.report_error(error);
                        } else {
                            self.notify(format!(
                                "otwieram zasób {}",
                                resource.effective_display_name()
                            ));
                        }
                    } else {
                        match open_network_resource(&resource) {
                            Ok(()) => self.notify(format!(
                                "otwieram zasób {}",
                                resource.effective_display_name()
                            )),
                            Err(error) => self.report_error(error),
                        }
                    }
                }
            }
            EntryKind::FavoriteDirectoriesRoot => {
                self.panels[panel_index].search_state = None;
                self.panels[panel_index].location = PanelLocation::FavoriteDirectories;
                if let Err(error) = self.refresh_and_keep(panel_index, None) {
                    self.report_error(error);
                }
            }
            EntryKind::FavoriteFilesRoot => {
                self.panels[panel_index].search_state = None;
                self.panels[panel_index].location = PanelLocation::FavoriteFiles;
                if let Err(error) = self.refresh_and_keep(panel_index, None) {
                    self.report_error(error);
                }
            }
            EntryKind::Parent => {
                if let Some(path) = entry.path {
                    let child = match &self.panels[panel_index].location {
                        PanelLocation::Filesystem(current) => Some(current.clone()),
                        PanelLocation::Remote(current) => Some(current.path.clone()),
                        PanelLocation::Archive(current) => Some(current.inside_path.clone()),
                        PanelLocation::Drives
                        | PanelLocation::NetworkResources
                        | PanelLocation::FavoriteDirectories
                        | PanelLocation::FavoriteFiles => None,
                    };
                    match &self.panels[panel_index].location {
                        PanelLocation::Remote(current) => {
                            let next_remote =
                                if current.sftp_privilege_mode != SftpPrivilegeMode::Normal {
                                    current.downgraded_sftp_location(path.clone())
                                } else {
                                    RemoteLocation {
                                        resource: current.resource.clone(),
                                        original_resource: current.original_resource.clone(),
                                        path: path.clone(),
                                        sftp_privilege_mode: current.sftp_privilege_mode,
                                    }
                                };
                            self.panels[panel_index].location = PanelLocation::Remote(next_remote);
                        }
                        PanelLocation::Archive(current) => {
                            if path == current.archive_path {
                                let archive_path = current.archive_path.clone();
                                let parent = archive_path
                                    .parent()
                                    .map(Path::to_path_buf)
                                    .unwrap_or_else(|| PathBuf::from("."));
                                self.panels[panel_index].location =
                                    PanelLocation::Filesystem(parent.clone());
                                self.panels[panel_index].last_filesystem_path = parent;
                            } else {
                                self.panels[panel_index].location =
                                    PanelLocation::Archive(ArchiveLocation {
                                        archive_path: current.archive_path.clone(),
                                        inside_path: normalize_archive_inner_path(&path),
                                    });
                            }
                        }
                        _ => {
                            self.panels[panel_index].location =
                                PanelLocation::Filesystem(path.clone());
                            self.panels[panel_index].last_filesystem_path = path.clone();
                        }
                    }
                    if let Err(error) = self.refresh_and_keep(panel_index, child.as_deref()) {
                        self.report_error(error);
                    }
                }
            }
            EntryKind::Directory | EntryKind::Drive => {
                if let Some(path) = entry.path {
                    self.enter_directory(panel_index, path);
                }
            }
            EntryKind::File => match self.panels[panel_index].location.clone() {
                PanelLocation::Remote(remote) => {
                    self.open_remote_file_or_link(panel_index, remote, entry);
                }
                PanelLocation::Archive(archive) => {
                    if let Some(path) = entry.path {
                        self.open_archive_file_in_background(archive, path, entry.name);
                    }
                }
                _ => self.open_local_file_or_link(panel_index, entry),
            },
            EntryKind::NetworkPlaceholder => {
                self.panels[panel_index].search_state = None;
                self.panels[panel_index].location = PanelLocation::NetworkResources;
                if let Err(error) = self.refresh_and_keep(panel_index, None) {
                    self.report_error(error);
                }
            }
        }
    }

    unsafe fn go_up(&mut self, panel_index: usize) {
        let location = self.panels[panel_index].location.clone();
        match location {
            PanelLocation::Drives => {
                let restore = self.panels[panel_index].last_filesystem_path.clone();
                self.panels[panel_index].location = PanelLocation::Filesystem(restore.clone());
                if let Err(error) = self.refresh_and_keep(panel_index, Some(restore.as_path())) {
                    self.report_error(error);
                }
            }
            PanelLocation::NetworkResources => {
                self.panels[panel_index].location = PanelLocation::Drives;
                if let Err(error) = self.refresh_and_keep(panel_index, None) {
                    self.report_error(error);
                }
            }
            PanelLocation::Remote(remote) => {
                if let Some(parent) = remote_parent_path(&remote.path) {
                    let next_remote = if remote.sftp_privilege_mode != SftpPrivilegeMode::Normal {
                        remote.downgraded_sftp_location(parent.clone())
                    } else {
                        RemoteLocation {
                            resource: remote.resource,
                            original_resource: remote.original_resource,
                            path: parent.clone(),
                            sftp_privilege_mode: remote.sftp_privilege_mode,
                        }
                    };
                    self.panels[panel_index].location = PanelLocation::Remote(next_remote);
                    if let Err(error) = self.refresh_and_keep(panel_index, Some(parent.as_path())) {
                        self.report_error(error);
                    }
                } else {
                    self.panels[panel_index].location = PanelLocation::NetworkResources;
                    if let Err(error) = self.refresh_and_keep(panel_index, None) {
                        self.report_error(error);
                    }
                }
            }
            PanelLocation::Archive(archive) => {
                if archive.inside_path.as_os_str().is_empty() {
                    let archive_path = archive.archive_path.clone();
                    let parent = archive_path
                        .parent()
                        .map(Path::to_path_buf)
                        .unwrap_or_else(|| PathBuf::from("."));
                    self.panels[panel_index].location = PanelLocation::Filesystem(parent.clone());
                    self.panels[panel_index].last_filesystem_path = parent;
                    if let Err(error) = self.refresh_and_keep(panel_index, Some(&archive_path)) {
                        self.report_error(error);
                    }
                } else {
                    let select = archive.inside_path.clone();
                    let parent = archive
                        .inside_path
                        .parent()
                        .map(Path::to_path_buf)
                        .unwrap_or_default();
                    self.panels[panel_index].location = PanelLocation::Archive(ArchiveLocation {
                        archive_path: archive.archive_path,
                        inside_path: parent,
                    });
                    if let Err(error) = self.refresh_and_keep(panel_index, Some(&select)) {
                        self.report_error(error);
                    }
                }
            }
            PanelLocation::FavoriteDirectories | PanelLocation::FavoriteFiles => {
                self.panels[panel_index].location = PanelLocation::Drives;
                if let Err(error) = self.refresh_and_keep(panel_index, None) {
                    self.report_error(error);
                }
            }
            PanelLocation::Filesystem(path) => {
                if let Some(parent) = path.parent() {
                    let select = path.clone();
                    self.panels[panel_index].location =
                        PanelLocation::Filesystem(parent.to_path_buf());
                    self.panels[panel_index].last_filesystem_path = parent.to_path_buf();
                    if let Err(error) = self.refresh_and_keep(panel_index, Some(select.as_path())) {
                        self.report_error(error);
                    }
                } else {
                    self.panels[panel_index].location = PanelLocation::Drives;
                    if let Err(error) = self.refresh_and_keep(panel_index, None) {
                        self.report_error(error);
                    }
                }
            }
        }
    }

    unsafe fn refresh_and_keep(
        &mut self,
        panel_index: usize,
        preferred_path: Option<&Path>,
    ) -> io::Result<()> {
        self.refresh_panel_after_mutation(panel_index, preferred_path.map(Path::to_path_buf))?;
        self.focus_panel(panel_index, false);
        self.notify(if self.panels[panel_index].is_search_active() {
            "odświeżam wyniki wyszukiwania"
        } else {
            "ładowanie katalogu"
        });
        Ok(())
    }

    unsafe fn toggle_mark(&mut self, panel_index: usize) {
        let Some(entry) = self.panels[panel_index].selected_entry().cloned() else {
            self.notify("brak elementów");
            return;
        };
        if !entry.is_markable() {
            self.notify(entry.accessible_label(false, self.view_options));
            return;
        }
        let Some(path) = entry.path.clone() else {
            self.notify(entry.accessible_label(false, self.view_options));
            return;
        };

        let panel = &mut self.panels[panel_index];
        let selected = panel.selected;
        if !panel.marked.insert(path.clone()) {
            panel.marked.remove(&path);
        }
        self.rebuild_panel(panel_index);
        self.update_listbox_view(self.panels[panel_index].list_hwnd, selected);
        self.focus_panel(panel_index, false);
        self.notify(
            self.panels[panel_index]
                .selected_entry()
                .map(|current| {
                    let marked = current
                        .path
                        .as_ref()
                        .map(|item| self.panels[panel_index].marked.contains(item))
                        .unwrap_or(false);
                    current.accessible_label(marked, self.view_options)
                })
                .unwrap_or_else(|| "brak elementów".to_string()),
        );
    }

    unsafe fn mark_all(&mut self, panel_index: usize) {
        let panel = &mut self.panels[panel_index];
        panel.marked.clear();
        for entry in &panel.entries {
            if entry.is_markable() {
                if let Some(path) = &entry.path {
                    panel.marked.insert(path.clone());
                }
            }
        }
        let count = panel.marked.len();
        self.rebuild_panel(panel_index);
        self.focus_panel(panel_index, false);
        self.notify(format!("zaznaczono {}", pluralized_elements(count)));
    }

    unsafe fn unmark_all(&mut self, panel_index: usize) {
        self.panels[panel_index].marked.clear();
        self.rebuild_panel(panel_index);
        self.focus_panel(panel_index, false);
        self.notify("odznaczono wszystkie elementy");
    }

    unsafe fn invert_marks(&mut self, panel_index: usize) {
        let panel = &mut self.panels[panel_index];
        let visible_paths = panel
            .entries
            .iter()
            .filter(|entry| entry.is_markable())
            .filter_map(|entry| entry.path.clone())
            .collect::<Vec<_>>();
        for path in visible_paths {
            if !panel.marked.insert(path.clone()) {
                panel.marked.remove(&path);
            }
        }
        let count = panel
            .entries
            .iter()
            .filter_map(|entry| entry.path.as_ref())
            .filter(|path| panel.marked.contains(*path))
            .count();
        self.rebuild_panel(panel_index);
        self.focus_panel(panel_index, false);
        self.notify(format!("zaznaczono {}", pluralized_elements(count)));
    }

    unsafe fn refresh_view_options(&mut self) -> io::Result<()> {
        self.create_menu();
        for panel_index in 0..2 {
            let preferred_path = self.panels[panel_index]
                .selected_entry()
                .and_then(|entry| entry.path.clone());
            self.refresh_panel_after_mutation(panel_index, preferred_path)?;
        }
        self.focus_panel(self.active_panel, false);
        self.notify("odświeżono widok");
        Ok(())
    }

    unsafe fn refresh_panel_after_mutation(
        &mut self,
        panel_index: usize,
        preferred_path: Option<PathBuf>,
    ) -> io::Result<()> {
        if let Some(search) = self.panels[panel_index].search_state.clone() {
            self.apply_search(
                panel_index,
                &search.pattern,
                search.recursive,
                preferred_path,
            )?;
        } else {
            self.start_panel_load(panel_index, preferred_path);
        }
        Ok(())
    }

    unsafe fn show_search_dialog_for_panel(&mut self, panel_index: usize) {
        let Some(_) = self.panels[panel_index].current_dir_owned() else {
            self.notify("wyszukiwanie działa tylko w katalogu");
            return;
        };
        let Some((pattern, recursive)) = ({
            self.notify("wyszukiwanie, pole edycji");
            show_search_dialog(self.hwnd, Some(&self.nvda))
        }) else {
            self.notify("anulowano");
            return;
        };
        match self.apply_search(panel_index, pattern.trim(), recursive, None) {
            Ok(()) => {}
            Err(error) => self.report_error(error),
        }
    }

    unsafe fn apply_search(
        &mut self,
        panel_index: usize,
        pattern: &str,
        recursive: bool,
        preferred_path: Option<PathBuf>,
    ) -> io::Result<()> {
        let Some(base_dir) = self.panels[panel_index].current_dir_owned() else {
            return Err(io::Error::other("wyszukiwanie działa tylko w katalogu"));
        };
        let pattern = pattern.trim();
        if pattern.is_empty() {
            return Err(io::Error::other(
                "wyrażenie do wyszukania nie może być puste",
            ));
        }
        let generation = {
            let panel = &mut self.panels[panel_index];
            panel.search_generation = panel.search_generation.wrapping_add(1);
            panel.search_state = Some(SearchState {
                pattern: pattern.to_string(),
                recursive,
            });
            panel.search_in_progress = true;
            panel.loading = false;
            panel.pending_select = None;
            panel.pending_key = None;
            panel.entries = vec![PanelEntry::search_loading()];
            panel.selected = 0;
            panel.search_generation
        };
        self.rebuild_panel(panel_index);
        self.focus_panel(panel_index, false);
        self.notify("trwa wyszukiwanie");
        UpdateWindow(self.hwnd);
        let sender = self.search_sender.clone();
        let hwnd = self.hwnd as usize;
        let view_options = self.view_options;
        let pattern = pattern.to_string();
        thread::spawn(move || {
            let result = compile_search_regex(&pattern)
                .and_then(|regex| search_entries(&base_dir, recursive, &regex, view_options));
            let message = match result {
                Ok(entries) => SearchMessage {
                    panel_index,
                    generation,
                    match_count: entries.len(),
                    entries,
                    preferred_path,
                    error: None,
                },
                Err(error) => SearchMessage {
                    panel_index,
                    generation,
                    match_count: 0,
                    entries: Vec::new(),
                    preferred_path,
                    error: Some(error.to_string()),
                },
            };
            let _ = sender.send(message);
            unsafe {
                SendMessageW(hwnd as HWND, WM_SEARCH_EVENT, 0, 0);
            }
        });
        Ok(())
    }

    unsafe fn exit_search_results(&mut self, panel_index: usize) {
        if !self.panels[panel_index].is_search_active() {
            return;
        }
        self.panels[panel_index].search_generation =
            self.panels[panel_index].search_generation.wrapping_add(1);
        self.panels[panel_index].search_in_progress = false;
        let preferred = self.panels[panel_index]
            .selected_entry()
            .and_then(|entry| entry.path.clone())
            .filter(|path| path.parent() == self.panels[panel_index].current_dir());
        self.panels[panel_index].search_state = None;
        if let Err(error) = self.refresh_and_keep(panel_index, preferred.as_deref()) {
            self.report_error(error);
        } else {
            self.notify("wyjście z wyników wyszukiwania");
        }
    }

    unsafe fn mark_by_extension(&mut self, panel_index: usize) {
        let Some(extension) = ({
            self.notify("zaznaczanie po rozszerzeniu, pole edycji");
            show_input_dialog(
                self.hwnd,
                "Zaznacz po rozszerzeniu",
                "Rozszerzenie:",
                "",
                Some(&self.nvda),
            )
        }) else {
            self.notify("anulowano");
            return;
        };
        let normalized = extension.trim().trim_start_matches('.').to_lowercase();
        if normalized.is_empty() {
            self.notify("rozszerzenie nie może być puste");
            return;
        }
        let panel = &mut self.panels[panel_index];
        let mut matched = 0usize;
        for entry in &panel.entries {
            if entry.kind != EntryKind::File {
                continue;
            }
            let Some(path) = &entry.path else {
                continue;
            };
            let entry_extension = path
                .extension()
                .map(|item| item.to_string_lossy().to_lowercase())
                .unwrap_or_default();
            if entry_extension == normalized {
                panel.marked.insert(path.clone());
                matched += 1;
            }
        }
        self.rebuild_panel(panel_index);
        self.focus_panel(panel_index, false);
        self.notify(format!("zaznaczono {}", pluralized_elements(matched)));
    }

    unsafe fn mark_by_name_pattern(&mut self, panel_index: usize) {
        let Some(pattern) = ({
            self.notify("zaznaczanie po nazwie, pole edycji");
            show_input_dialog(
                self.hwnd,
                "Zaznacz po nazwie",
                "Nazwa lub wyrażenie regularne:",
                "",
                Some(&self.nvda),
            )
        }) else {
            self.notify("anulowano");
            return;
        };
        let pattern = pattern.trim();
        if pattern.is_empty() {
            self.notify("nazwa nie może być pusta");
            return;
        }
        let regex = match RegexBuilder::new(pattern).case_insensitive(true).build() {
            Ok(regex) => regex,
            Err(error) => {
                self.report_error(format!("nieprawidłowe wyrażenie regularne: {error}"));
                return;
            }
        };
        let panel = &mut self.panels[panel_index];
        let mut matched = 0usize;
        for entry in &panel.entries {
            if !entry.is_markable() {
                continue;
            }
            let Some(path) = &entry.path else {
                continue;
            };
            if regex.is_match(&entry.name) {
                panel.marked.insert(path.clone());
                matched += 1;
            }
        }
        self.rebuild_panel(panel_index);
        self.focus_panel(panel_index, false);
        self.notify(format!("zaznaczono {}", pluralized_elements(matched)));
    }

    fn current_targets(&self, panel_index: usize) -> Result<Vec<PathBuf>, String> {
        let panel = &self.panels[panel_index];
        let mut items = panel
            .entries
            .iter()
            .filter_map(|entry| entry.path.as_ref())
            .filter(|path| panel.marked.contains(*path))
            .cloned()
            .collect::<Vec<_>>();
        if !items.is_empty() {
            items.sort();
            return Ok(items);
        }

        let Some(entry) = panel.selected_entry() else {
            return Err("brak elementów".to_string());
        };
        if !entry.is_operable() {
            return Err("wybierz plik lub katalog".to_string());
        }
        entry
            .path
            .clone()
            .map(|path| vec![path])
            .ok_or_else(|| "nie można wykonać operacji".to_string())
    }

    unsafe fn copy_selection_to_clipboard(&mut self, panel_index: usize, cut_mode: bool) {
        let targets = match self.current_targets(panel_index) {
            Ok(targets) => targets,
            Err(message) => {
                self.notify(message);
                return;
            }
        };
        let source_remote = self.remote_location_for_panel(panel_index);
        let source_entries = self.panels[panel_index].entries.clone();
        let requested_operation = if cut_mode {
            ClipboardOperation::Move
        } else {
            ClipboardOperation::Copy
        };

        let (paths_for_clipboard, system_operation) = if let Some(remote) = &source_remote {
            match self.stage_remote_targets_for_clipboard(remote, &targets) {
                Ok(paths) => {
                    if cut_mode {
                        self.notify_non_interrupting(
                            "zdalne wycięcie do schowka systemowego będzie wklejane jako kopia poza Amiga FM"
                                .to_string(),
                        );
                    }
                    (paths, ClipboardOperation::Copy)
                }
                Err(error) => {
                    self.report_error(error);
                    return;
                }
            }
        } else {
            (targets.clone(), requested_operation)
        };

        match write_file_paths_to_clipboard(self.hwnd, &paths_for_clipboard, system_operation) {
            Ok(sequence) => {
                self.clear_internal_clipboard();
                self.clipboard = Some(AppClipboardState {
                    operation: requested_operation,
                    source_panel_index: Some(panel_index),
                    source_remote,
                    source_entries,
                    items: targets,
                    staged_paths: paths_for_clipboard,
                    clipboard_sequence: sequence,
                });
                self.notify(if cut_mode {
                    "wycięto do schowka"
                } else {
                    "skopiowano do schowka"
                });
            }
            Err(error) => self.report_error(error),
        }
    }

    unsafe fn paste_clipboard_into_panel(&mut self, panel_index: usize) {
        let current_sequence = GetClipboardSequenceNumber();
        let clipboard = self
            .clipboard
            .clone()
            .filter(|clipboard| clipboard.clipboard_sequence == current_sequence);
        let using_internal_clipboard = clipboard.is_some();

        let (operation, source_panel_index, source_remote, source_entries, targets) =
            if let Some(clipboard) = clipboard {
                (
                    clipboard.operation,
                    clipboard.source_panel_index,
                    clipboard.source_remote,
                    clipboard.source_entries,
                    clipboard.items,
                )
            } else {
                let Some((paths, operation)) = (match read_file_paths_from_clipboard(self.hwnd) {
                    Ok(value) => value,
                    Err(error) => {
                        self.report_error(error);
                        return;
                    }
                }) else {
                    self.notify("schowek nie zawiera plików ani katalogów");
                    return;
                };
                (operation, None, None, Vec::new(), paths)
            };

        if targets.is_empty() {
            self.notify("schowek nie zawiera plików ani katalogów");
            return;
        }

        self.transfer_targets(
            source_panel_index,
            panel_index,
            source_remote,
            source_entries,
            targets,
            operation == ClipboardOperation::Move,
            false,
        );
        if using_internal_clipboard && operation == ClipboardOperation::Move {
            self.clear_internal_clipboard();
        }
    }

    fn current_single_target(&self, panel_index: usize) -> Result<PathBuf, String> {
        let targets = self.current_targets(panel_index)?;
        match targets.as_slice() {
            [single] => Ok(single.clone()),
            [] => Err("brak elementów".to_string()),
            _ => Err("ta funkcja działa dla jednego elementu".to_string()),
        }
    }

    fn remote_location_for_panel(&self, panel_index: usize) -> Option<RemoteLocation> {
        match &self.panels[panel_index].location {
            PanelLocation::Remote(remote) => Some(remote.clone()),
            _ => None,
        }
    }

    fn clear_internal_clipboard(&mut self) {
        let Some(clipboard) = self.clipboard.take() else {
            return;
        };
        if clipboard.source_remote.is_none() {
            return;
        }

        let mut staging_roots = HashSet::new();
        for path in clipboard.staged_paths {
            if let Some(parent) = path.parent() {
                staging_roots.insert(parent.to_path_buf());
            }
        }
        for root in staging_roots {
            let _ = fs::remove_dir_all(root);
        }
    }

    unsafe fn stage_remote_targets_for_clipboard(
        &mut self,
        remote: &RemoteLocation,
        targets: &[PathBuf],
    ) -> io::Result<Vec<PathBuf>> {
        let staging_root = unique_temp_file_path("AmigaFmNativeClipboard", "staging")?;
        fs::create_dir_all(&staging_root)?;

        let mut used_names = HashSet::new();
        let staged_targets = targets
            .iter()
            .map(|target| unique_staging_destination(&staging_root, target, &mut used_names))
            .collect::<Vec<_>>();

        let mut counts = ItemCounts::default();
        let mut target_summaries = Vec::new();
        for target in targets {
            counts.add(summarize_remote_path(remote, target)?);
            target_summaries.push((target.clone(), summarize_remote_path(remote, target)?));
        }

        let mut active_remote = remote.clone();
        loop {
            let worker_remote = active_remote.clone();
            let worker_staged = staged_targets.clone();
            let worker_summaries = target_summaries.clone();
            let outcome = run_progress_dialog(
                self.hwnd,
                "Przygotowanie schowka",
                build_progress_lines(
                    "przygotowanie schowka",
                    "Przygotowanie kopiowanych elementów.",
                    None,
                    Some(&staging_root),
                    ItemCounts::default(),
                    counts,
                    Duration::ZERO,
                ),
                Some(&self.nvda),
                ProgressDialogOptions {
                    auto_close_on_success: true,
                    auto_close_on_retryable_error: true,
                },
                move |progress_hwnd, sender, cancel_flag, conflict_receiver| {
                    let mut progress = ProgressReporter::new(
                        sender,
                        progress_hwnd,
                        cancel_flag,
                        "przygotowanie schowka",
                        counts,
                        conflict_receiver,
                    );
                    for ((target, summary), destination) in
                        worker_summaries.iter().zip(worker_staged.iter())
                    {
                        let result = if worker_remote.uses_sftp_sudo() {
                            copy_remote_to_local_via_sudo(
                                &worker_remote,
                                target,
                                destination,
                                *summary,
                                &mut progress,
                            )
                        } else {
                            let mut client = match connect_remote_client(&worker_remote.resource) {
                                Ok(client) => client,
                                Err(error) => {
                                    progress.finish(
                                        &format!("Błąd: {error}"),
                                        format!("przygotowanie schowka zakończone błędem: {error}"),
                                        WorkerOutcome::Error(error.to_string()),
                                    );
                                    return;
                                }
                            };
                            let result = copy_remote_to_local_with_client(
                                &mut client,
                                target,
                                destination,
                                *summary,
                                &mut progress,
                            );
                            client.disconnect();
                            result
                        };
                        match result {
                            Ok(OperationResult::Done) | Ok(OperationResult::Skipped) => {}
                            Ok(OperationResult::Canceled) => {
                                progress.finish(
                                    "Przygotowanie schowka anulowane.",
                                    "anulowano".to_string(),
                                    WorkerOutcome::Canceled,
                                );
                                return;
                            }
                            Err(error) => {
                                progress.finish(
                                    &format!("Błąd: {error}"),
                                    format!("przygotowanie schowka zakończone błędem: {error}"),
                                    WorkerOutcome::Error(error.to_string()),
                                );
                                return;
                            }
                        }
                    }
                    progress.finish(
                        "Przygotowanie schowka ukończone.",
                        "schowek przygotowany".to_string(),
                        WorkerOutcome::Success,
                    );
                },
            )?;

            match outcome {
                WorkerOutcome::Success => return Ok(staged_targets),
                WorkerOutcome::Canceled => {
                    return Err(io::Error::new(io::ErrorKind::Interrupted, "anulowano"));
                }
                WorkerOutcome::Error(message) => {
                    if let Some(candidate) = self.next_sftp_retry_remote(&active_remote, &message) {
                        active_remote = candidate;
                        continue;
                    }
                    return Err(io::Error::other(message));
                }
            }
        }
    }

    unsafe fn transfer_targets(
        &mut self,
        source_panel_index: Option<usize>,
        destination_panel_index: usize,
        source_remote: Option<RemoteLocation>,
        _source_entries: Vec<PanelEntry>,
        targets: Vec<PathBuf>,
        move_mode: bool,
        show_prompt: bool,
    ) {
        let destination_remote = self.remote_location_for_panel(destination_panel_index);
        let mut source_sftp_fallback: Option<RemoteLocation> = None;
        let mut destination_sftp_fallback: Option<RemoteLocation> = None;
        let destination_dir = if let Some(remote) = &destination_remote {
            remote.path.clone()
        } else {
            let Some(destination_dir) = self.panels[destination_panel_index].current_dir_owned()
            else {
                self.notify("panel docelowy musi być katalogiem");
                return;
            };
            destination_dir
        };
        let destination_label = if let Some(remote) = &destination_remote {
            format!(
                "{}{}",
                remote.resource.effective_display_name(),
                remote_display_suffix(&remote.path)
            )
        } else {
            destination_dir.display().to_string()
        };

        let mut counts = ItemCounts::default();
        let mut target_summaries = Vec::new();
        for target in &targets {
            loop {
                let summary = if let Some(remote) = &source_remote {
                    summarize_remote_path(remote, target)
                } else {
                    summarize_path(target)
                };
                match summary {
                    Ok(summary) => {
                        counts.add(summary);
                        target_summaries.push((target.clone(), summary));
                        break;
                    }
                    Err(error) if is_sftp_retry_error(&error) => {
                        let fallback = if let Some(fallback) = source_sftp_fallback.clone() {
                            Some(fallback)
                        } else {
                            source_remote.as_ref().and_then(|remote| {
                                self.next_sftp_retry_remote(remote, &error.to_string())
                            })
                        };
                        if let Some(candidate) = fallback {
                            match summarize_remote_path(&candidate, target) {
                                Ok(summary) => {
                                    source_sftp_fallback = Some(candidate);
                                    counts.add(summary);
                                    target_summaries.push((target.clone(), summary));
                                    break;
                                }
                                Err(error) if is_sftp_retry_error(&error) => {
                                    if let Some(next_candidate) =
                                        self.next_sftp_retry_remote(&candidate, &error.to_string())
                                    {
                                        source_sftp_fallback = Some(next_candidate);
                                        continue;
                                    }
                                    self.report_error(error);
                                    return;
                                }
                                Err(error) => {
                                    self.report_error(error);
                                    return;
                                }
                            }
                        }
                        self.report_error(error);
                        return;
                    }
                    Err(error) => {
                        self.report_error(error);
                        return;
                    }
                }
            }
        }

        if show_prompt {
            let accepted = show_operation_prompt(
                self.hwnd,
                if move_mode {
                    "Przenoszenie"
                } else {
                    "Kopiowanie"
                },
                vec![
                    format!(
                        "Operacja: {}",
                        if move_mode {
                            "przenoszenie"
                        } else {
                            "kopiowanie"
                        }
                    ),
                    format!("Elementy: {}", summarize_targets(&targets)),
                    format!("Cel: {destination_label}"),
                    format!("Obiekty: {}", counts_description(counts)),
                    format!("Rozmiar: {}", format_bytes(counts.bytes)),
                ],
                Some(&self.nvda),
            );
            if !accepted {
                self.notify("anulowano");
                return;
            }
        }

        if source_remote.is_none() && destination_remote.is_none() {
            for target in &targets {
                if let Err(message) = validate_destination(target, &destination_dir) {
                    self.notify(message);
                    return;
                }
            }
        }

        let focus_target = targets
            .iter()
            .find_map(|target| target.file_name().map(|name| destination_dir.join(name)));
        let initial_lines = build_progress_lines(
            if move_mode {
                "przenoszenie"
            } else {
                "kopiowanie"
            },
            "Przygotowanie operacji.",
            None,
            Some(&destination_dir),
            ItemCounts::default(),
            counts,
            Duration::default(),
        );
        let outcome = loop {
            let progress_source_remote = source_remote.clone();
            let progress_destination_remote = destination_remote.clone();
            let progress_source_sftp_fallback = source_sftp_fallback.clone();
            let progress_destination_sftp_fallback = destination_sftp_fallback.clone();
            let progress_targets = target_summaries.clone();
            let progress_destination = destination_dir.clone();
            let outcome = match run_progress_dialog(
                self.hwnd,
                if move_mode {
                    "Postęp przenoszenia"
                } else {
                    "Postęp kopiowania"
                },
                initial_lines.clone(),
                Some(&self.nvda),
                ProgressDialogOptions {
                    auto_close_on_success: false,
                    auto_close_on_retryable_error: true,
                },
                move |progress_hwnd, sender, cancel_flag, conflict_receiver| {
                    let mut progress = ProgressReporter::new(
                        sender,
                        progress_hwnd,
                        cancel_flag,
                        if move_mode {
                            "przenoszenie"
                        } else {
                            "kopiowanie"
                        },
                        counts,
                        conflict_receiver,
                    );

                    for (target, summary) in &progress_targets {
                        let Some(file_name) = target.file_name() else {
                            continue;
                        };
                        let destination = progress_destination.join(file_name);
                        let result = match (&progress_source_remote, &progress_destination_remote) {
                            (None, None) => {
                                if move_mode {
                                    move_item(target, &destination, *summary, &mut progress)
                                } else {
                                    copy_item(target, &destination, *summary, &mut progress)
                                }
                            }
                            (None, Some(remote)) => {
                                let result: io::Result<OperationResult> = (|| {
                                    let result = copy_local_to_remote_with_sftp_fallback(
                                        remote,
                                        progress_destination_sftp_fallback.as_ref(),
                                        target,
                                        &destination,
                                        *summary,
                                        &mut progress,
                                    );
                                    if matches!(result, Ok(OperationResult::Done)) && move_mode {
                                        delete_path_after_move(target)?;
                                    }
                                    result
                                })(
                                );
                                result
                            }
                            (Some(remote), None) => {
                                let result: io::Result<OperationResult> = (|| {
                                    let result = copy_remote_to_local_with_sftp_fallback(
                                        remote,
                                        progress_source_sftp_fallback.as_ref(),
                                        target,
                                        &destination,
                                        *summary,
                                        &mut progress,
                                    );
                                    if matches!(result, Ok(OperationResult::Done)) && move_mode {
                                        delete_remote_path_after_move_with_sftp_fallback(
                                            remote,
                                            progress_source_sftp_fallback.as_ref(),
                                            target,
                                        )?;
                                    }
                                    result
                                })(
                                );
                                result
                            }
                            (Some(source_remote), Some(destination_remote)) => {
                                let same_resource = source_remote.resource.stable_key()
                                    == destination_remote.resource.stable_key();
                                let result: io::Result<OperationResult> = (|| {
                                    if move_mode && same_resource {
                                        return move_remote_with_sftp_fallback(
                                            source_remote,
                                            progress_source_sftp_fallback.as_ref(),
                                            target,
                                            &destination,
                                            *summary,
                                            &mut progress,
                                        );
                                    }

                                    let result = copy_remote_to_remote_with_sftp_fallback(
                                        source_remote,
                                        destination_remote,
                                        progress_source_sftp_fallback.as_ref(),
                                        progress_destination_sftp_fallback.as_ref(),
                                        target,
                                        &destination,
                                        *summary,
                                        &mut progress,
                                    );
                                    if matches!(result, Ok(OperationResult::Done)) && move_mode {
                                        delete_remote_path_after_move_with_sftp_fallback(
                                            source_remote,
                                            progress_source_sftp_fallback.as_ref(),
                                            target,
                                        )?;
                                    }
                                    result
                                })(
                                );
                                result
                            }
                        };
                        match result {
                            Ok(OperationResult::Done) => {}
                            Ok(OperationResult::Skipped) => {
                                progress.add_counts(
                                    *summary,
                                    "Element pominięto.",
                                    Some(target),
                                    Some(&destination),
                                );
                            }
                            Ok(OperationResult::Canceled) => {
                                progress.finish(
                                    "Operacja anulowana.",
                                    "anulowano".to_string(),
                                    WorkerOutcome::Canceled,
                                );
                                return;
                            }
                            Err(error) => {
                                progress.finish(
                                    &format!("Błąd: {error}"),
                                    format!("operacja zakończona błędem: {error}"),
                                    WorkerOutcome::Error(error.to_string()),
                                );
                                return;
                            }
                        }
                    }

                    progress.finish(
                        if move_mode {
                            "Przenoszenie ukończone."
                        } else {
                            "Kopiowanie ukończone."
                        },
                        if move_mode {
                            "przenoszenie ukończone".to_string()
                        } else {
                            "kopiowanie ukończone".to_string()
                        },
                        WorkerOutcome::Success,
                    );
                },
            ) {
                Ok(outcome) => outcome,
                Err(error) => {
                    self.report_error(error);
                    return;
                }
            };

            let mut retry_remote = false;
            if let WorkerOutcome::Error(message) = &outcome
                && is_sftp_retry_message(message)
            {
                if is_sftp_source_retry_message(message) {
                    if let Some(remote) = source_sftp_fallback.clone()
                        && let Some(candidate) = self.next_sftp_retry_remote(&remote, message)
                    {
                        source_sftp_fallback = Some(candidate);
                        retry_remote = true;
                    }
                    if !retry_remote
                        && let Some(remote) = source_remote.clone()
                        && let Some(candidate) = self.next_sftp_retry_remote(&remote, message)
                    {
                        source_sftp_fallback = Some(candidate);
                        retry_remote = true;
                    }
                } else if is_sftp_destination_retry_message(message) {
                    if let Some(remote) = destination_sftp_fallback.clone()
                        && let Some(candidate) = self.next_sftp_retry_remote(&remote, message)
                    {
                        destination_sftp_fallback = Some(candidate);
                        retry_remote = true;
                    }
                    if !retry_remote
                        && let Some(remote) = destination_remote.clone()
                        && let Some(candidate) = self.next_sftp_retry_remote(&remote, message)
                    {
                        destination_sftp_fallback = Some(candidate);
                        retry_remote = true;
                    }
                } else {
                    if let Some(remote) = source_sftp_fallback.clone()
                        && let Some(candidate) = self.next_sftp_retry_remote(&remote, message)
                    {
                        source_sftp_fallback = Some(candidate);
                        retry_remote = true;
                    }
                    if !retry_remote
                        && source_sftp_fallback.is_some()
                        && let Some(remote) = destination_remote.clone()
                        && let Some(candidate) = self.next_sftp_retry_remote(&remote, message)
                    {
                        destination_sftp_fallback = Some(candidate);
                        retry_remote = true;
                    }
                    if !retry_remote
                        && let Some(remote) = destination_sftp_fallback.clone()
                        && let Some(candidate) = self.next_sftp_retry_remote(&remote, message)
                    {
                        destination_sftp_fallback = Some(candidate);
                        retry_remote = true;
                    }
                    if !retry_remote
                        && let Some(remote) = source_remote.clone()
                        && let Some(candidate) = self.next_sftp_retry_remote(&remote, message)
                    {
                        source_sftp_fallback = Some(candidate);
                        retry_remote = true;
                    }
                    if !retry_remote
                        && let Some(remote) = destination_remote.clone()
                        && let Some(candidate) = self.next_sftp_retry_remote(&remote, message)
                    {
                        destination_sftp_fallback = Some(candidate);
                        retry_remote = true;
                    }
                }
            }
            if retry_remote {
                continue;
            }
            break outcome;
        };

        match outcome {
            WorkerOutcome::Success => {}
            WorkerOutcome::Canceled => {
                self.notify("anulowano");
                return;
            }
            WorkerOutcome::Error(message)
                if source_remote.is_none()
                    && destination_remote.is_none()
                    && is_permission_denied_message(&message) =>
            {
                if !self.confirm_elevated_operation(
                    if move_mode {
                        "przenoszenie"
                    } else {
                        "kopiowanie"
                    },
                    vec![
                        format!("Elementy: {}", summarize_targets(&targets)),
                        format!("Cel: {destination_label}"),
                        format!("Obiekty: {}", counts_description(counts)),
                    ],
                ) {
                    self.notify("anulowano");
                    return;
                }
                match run_elevated_local_operation(ElevatedLocalOperation::CopyMove {
                    sources: targets.clone(),
                    destination_dir: destination_dir.clone(),
                    move_mode,
                }) {
                    Ok(()) => {}
                    Err(error) if is_uac_canceled_error(&error) => {
                        self.notify("anulowano");
                        return;
                    }
                    Err(error) => {
                        self.report_error(error);
                        return;
                    }
                }
            }
            WorkerOutcome::Error(message) => {
                self.report_error(message);
                return;
            }
        }

        if let Some(source_panel_index) = source_panel_index {
            self.panels[source_panel_index]
                .marked
                .retain(|path| !targets.iter().any(|target| path == target));
        }

        if let Some(source_panel_index) = source_panel_index
            && source_panel_index != destination_panel_index
        {
            if let Err(error) = self.refresh_panel_after_mutation(source_panel_index, None) {
                self.report_error(error);
                return;
            }
        }
        if let Err(error) =
            self.refresh_panel_after_mutation(destination_panel_index, focus_target.clone())
        {
            self.report_error(error);
            return;
        }
        self.active_panel = destination_panel_index;
        self.focus_panel(destination_panel_index, false);
        self.notify_non_interrupting(format!(
            "{}, {}",
            if move_mode {
                "przenoszenie ukończone"
            } else {
                "kopiowanie ukończone"
            },
            counts_description(counts)
        ));
    }

    unsafe fn refresh_favorite_panels(&mut self) {
        for panel_index in 0..2 {
            if matches!(
                self.panels[panel_index].location,
                PanelLocation::FavoriteDirectories | PanelLocation::FavoriteFiles
            ) {
                let _ = self.refresh_panel_after_mutation(panel_index, None);
            }
        }
    }

    unsafe fn refresh_network_panels(&mut self) {
        for panel_index in 0..2 {
            if matches!(
                self.panels[panel_index].location,
                PanelLocation::NetworkResources | PanelLocation::Remote(_)
            ) {
                let _ = self.refresh_panel_after_mutation(panel_index, None);
            }
        }
    }

    unsafe fn add_to_favorites(&mut self, panel_index: usize) {
        if self.remote_location_for_panel(panel_index).is_some() {
            self.notify("ulubione działają tylko dla lokalnych elementów");
            return;
        }
        let targets = match self.current_targets(panel_index) {
            Ok(targets) => targets,
            Err(message) => {
                self.notify(message);
                return;
            }
        };

        let mut added_directories = 0usize;
        let mut added_files = 0usize;
        for target in targets {
            let metadata = match fs::symlink_metadata(&target) {
                Ok(metadata) => metadata,
                Err(_) => continue,
            };
            if metadata.is_dir() {
                if !self.settings.favorite_directories.contains(&target) {
                    self.settings.favorite_directories.push(target);
                    added_directories += 1;
                }
            } else if metadata.is_file() || metadata.file_type().is_symlink() {
                if !self.settings.favorite_files.contains(&target) {
                    self.settings.favorite_files.push(target);
                    added_files += 1;
                }
            }
        }

        self.settings
            .favorite_directories
            .sort_by_key(|path| path.to_string_lossy().to_lowercase());
        self.settings
            .favorite_files
            .sort_by_key(|path| path.to_string_lossy().to_lowercase());

        self.save_settings_or_report();
        self.refresh_favorite_panels();

        let summary = match (added_directories, added_files) {
            (0, 0) => "element już jest w ulubionych".to_string(),
            (directories, 0) => {
                format!(
                    "dodano do ulubionych {}",
                    pluralized_directories(directories)
                )
            }
            (0, files) => format!("dodano do ulubionych {}", pluralized_files(files)),
            (directories, files) => format!(
                "dodano do ulubionych {}, {}",
                pluralized_directories(directories),
                pluralized_files(files)
            ),
        };
        self.notify(summary);
    }

    unsafe fn add_network_resource(&mut self, resource: NetworkResource) {
        let key = resource.stable_key();
        if self
            .settings
            .network_resources
            .iter()
            .any(|existing| existing.stable_key() == key)
        {
            self.notify("połączenie sieciowe już istnieje");
            return;
        }
        self.settings.network_resources.push(resource.clone());
        self.settings.network_resources.sort_by(|left, right| {
            left.effective_display_name()
                .to_lowercase()
                .cmp(&right.effective_display_name().to_lowercase())
        });
        self.save_settings_or_report();
        self.refresh_network_panels();
        self.notify(format!(
            "dodano połączenie sieciowe {}",
            resource.effective_display_name()
        ));
    }

    fn remembered_sftp_root_password(&self, resource: &NetworkResource) -> Option<String> {
        self.settings
            .network_resources
            .iter()
            .find(|existing| existing.stable_key() == resource.stable_key())
            .map(|existing| existing.root_password.clone())
            .filter(|password| !password.is_empty())
            .or_else(|| {
                (!resource.root_password.is_empty()).then(|| resource.root_password.clone())
            })
    }

    fn remembered_sftp_sudo_password(&self, resource: &NetworkResource) -> Option<String> {
        self.settings
            .network_resources
            .iter()
            .find(|existing| existing.stable_key() == resource.stable_key())
            .map(|existing| existing.sudo_password.clone())
            .filter(|password| !password.is_empty())
            .or_else(|| {
                (!resource.sudo_password.is_empty()).then(|| resource.sudo_password.clone())
            })
    }

    unsafe fn remember_sftp_root_password(&mut self, resource: &NetworkResource, password: &str) {
        let key = resource.stable_key();
        for existing in &mut self.settings.network_resources {
            if existing.stable_key() == key {
                existing.root_password = password.to_string();
            }
        }
        for panel in &mut self.panels {
            if let PanelLocation::Remote(remote) = &mut panel.location
                && remote.resource.stable_key() == key
            {
                remote.resource.root_password = password.to_string();
                remote.original_resource.root_password = password.to_string();
            }
        }
        self.save_settings_or_report();
    }

    unsafe fn remember_sftp_sudo_password(&mut self, resource: &NetworkResource, password: &str) {
        let key = resource.stable_key();
        for existing in &mut self.settings.network_resources {
            if existing.stable_key() == key {
                existing.sudo_password = password.to_string();
            }
        }
        for panel in &mut self.panels {
            if let PanelLocation::Remote(remote) = &mut panel.location
                && remote.resource.stable_key() == key
            {
                remote.resource.sudo_password = password.to_string();
                remote.original_resource.sudo_password = password.to_string();
            }
        }
        self.save_settings_or_report();
    }

    unsafe fn clear_sftp_root_password(&mut self, resource: &NetworkResource) {
        let key = resource.stable_key();
        let mut changed = false;
        for existing in &mut self.settings.network_resources {
            if existing.stable_key() == key && !existing.root_password.is_empty() {
                existing.root_password.clear();
                changed = true;
            }
        }
        for panel in &mut self.panels {
            if let PanelLocation::Remote(remote) = &mut panel.location
                && remote.resource.stable_key() == key
                && !remote.resource.root_password.is_empty()
            {
                remote.resource.root_password.clear();
                remote.original_resource.root_password.clear();
            }
        }
        if changed {
            self.save_settings_or_report();
        }
    }

    unsafe fn clear_sftp_sudo_password(&mut self, resource: &NetworkResource) {
        let key = resource.stable_key();
        let mut changed = false;
        for existing in &mut self.settings.network_resources {
            if existing.stable_key() == key && !existing.sudo_password.is_empty() {
                existing.sudo_password.clear();
                changed = true;
            }
        }
        for panel in &mut self.panels {
            if let PanelLocation::Remote(remote) = &mut panel.location
                && remote.resource.stable_key() == key
                && !remote.resource.sudo_password.is_empty()
            {
                remote.resource.sudo_password.clear();
                remote.original_resource.sudo_password.clear();
            }
        }
        if changed {
            self.save_settings_or_report();
        }
    }

    unsafe fn resolve_sftp_root_resource(
        &mut self,
        resource: &NetworkResource,
        force_prompt: bool,
    ) -> Option<NetworkResource> {
        if resource.protocol != NetworkProtocol::Sftp {
            return None;
        }

        let password = if !force_prompt {
            if let Some(password) = self.remembered_sftp_root_password(resource) {
                password
            } else {
                self.notify("hasło roota sftp, pole edycji");
                let password = show_input_dialog(
                    self.hwnd,
                    "Uprawnienia root SFTP",
                    "Hasło roota:",
                    "",
                    Some(&self.nvda),
                )?;
                let trimmed = password.trim().to_string();
                if trimmed.is_empty() {
                    self.notify("hasło roota nie może być puste");
                    return None;
                }
                let remember = show_choice_prompt(
                    self.hwnd,
                    "Zapamiętać hasło roota",
                    vec![
                        "Czy zapamiętać hasło roota dla tego połączenia SFTP?".to_string(),
                        resource.summary_line(),
                    ],
                    vec![
                        DialogButton {
                            id: ID_DIALOG_YES,
                            label: "Tak",
                            is_default: true,
                        },
                        DialogButton {
                            id: ID_DIALOG_NO,
                            label: "Nie",
                            is_default: false,
                        },
                    ],
                    Some(&self.nvda),
                ) == ID_DIALOG_YES;
                if remember {
                    self.remember_sftp_root_password(resource, &trimmed);
                }
                trimmed
            }
        } else {
            self.notify("zapisane hasło roota jest nieaktualne, podaj nowe hasło");
            self.clear_sftp_root_password(resource);
            self.notify("hasło roota sftp, pole edycji");
            let password = show_input_dialog(
                self.hwnd,
                "Uprawnienia root SFTP",
                "Hasło roota:",
                "",
                Some(&self.nvda),
            )?;
            let trimmed = password.trim().to_string();
            if trimmed.is_empty() {
                self.notify("hasło roota nie może być puste");
                return None;
            }
            let remember = show_choice_prompt(
                self.hwnd,
                "Zapamiętać hasło roota",
                vec![
                    "Czy zapamiętać nowe hasło roota dla tego połączenia SFTP?".to_string(),
                    resource.summary_line(),
                ],
                vec![
                    DialogButton {
                        id: ID_DIALOG_YES,
                        label: "Tak",
                        is_default: true,
                    },
                    DialogButton {
                        id: ID_DIALOG_NO,
                        label: "Nie",
                        is_default: false,
                    },
                ],
                Some(&self.nvda),
            ) == ID_DIALOG_YES;
            if remember {
                self.remember_sftp_root_password(resource, &trimmed);
            }
            trimmed
        };

        let mut elevated = resource.clone();
        elevated.username = "root".to_string();
        elevated.password = password.clone();
        elevated.root_password = password;
        elevated.anonymous = false;
        elevated.ssh_key.clear();
        Some(elevated)
    }

    unsafe fn resolve_sftp_sudo_resource(
        &mut self,
        resource: &NetworkResource,
        force_prompt: bool,
    ) -> Option<NetworkResource> {
        if resource.protocol != NetworkProtocol::Sftp
            || resource.username.trim().is_empty()
            || resource.username.eq_ignore_ascii_case("root")
        {
            return None;
        }

        let password = if !force_prompt {
            if let Some(password) = self.remembered_sftp_sudo_password(resource) {
                password
            } else {
                let accepted = show_choice_prompt(
                    self.hwnd,
                    "Uprawnienia sudo SFTP",
                    vec![
                        "Logowanie jako root nie powiodło się.".to_string(),
                        format!(
                            "Czy spróbować wykonać operację przez sudo dla użytkownika {}?",
                            resource.username.trim()
                        ),
                        resource.summary_line(),
                    ],
                    vec![
                        DialogButton {
                            id: ID_DIALOG_YES,
                            label: "Tak",
                            is_default: true,
                        },
                        DialogButton {
                            id: ID_DIALOG_NO,
                            label: "Nie",
                            is_default: false,
                        },
                    ],
                    Some(&self.nvda),
                ) == ID_DIALOG_YES;
                if !accepted {
                    return None;
                }
                self.notify("hasło sudo sftp, pole edycji");
                let password = show_input_dialog(
                    self.hwnd,
                    "Uprawnienia sudo SFTP",
                    "Hasło sudo:",
                    "",
                    Some(&self.nvda),
                )?;
                let trimmed = password.trim().to_string();
                if trimmed.is_empty() {
                    self.notify("hasło sudo nie może być puste");
                    return None;
                }
                let remember = show_choice_prompt(
                    self.hwnd,
                    "Zapamiętać hasło sudo",
                    vec![
                        "Czy zapamiętać hasło sudo dla tego połączenia SFTP?".to_string(),
                        resource.summary_line(),
                    ],
                    vec![
                        DialogButton {
                            id: ID_DIALOG_YES,
                            label: "Tak",
                            is_default: true,
                        },
                        DialogButton {
                            id: ID_DIALOG_NO,
                            label: "Nie",
                            is_default: false,
                        },
                    ],
                    Some(&self.nvda),
                ) == ID_DIALOG_YES;
                if remember {
                    self.remember_sftp_sudo_password(resource, &trimmed);
                }
                trimmed
            }
        } else {
            self.notify("zapisane hasło sudo jest nieaktualne, podaj nowe hasło");
            self.clear_sftp_sudo_password(resource);
            self.notify("hasło sudo sftp, pole edycji");
            let password = show_input_dialog(
                self.hwnd,
                "Uprawnienia sudo SFTP",
                "Hasło sudo:",
                "",
                Some(&self.nvda),
            )?;
            let trimmed = password.trim().to_string();
            if trimmed.is_empty() {
                self.notify("hasło sudo nie może być puste");
                return None;
            }
            let remember = show_choice_prompt(
                self.hwnd,
                "Zapamiętać hasło sudo",
                vec![
                    "Czy zapamiętać nowe hasło sudo dla tego połączenia SFTP?".to_string(),
                    resource.summary_line(),
                ],
                vec![
                    DialogButton {
                        id: ID_DIALOG_YES,
                        label: "Tak",
                        is_default: true,
                    },
                    DialogButton {
                        id: ID_DIALOG_NO,
                        label: "Nie",
                        is_default: false,
                    },
                ],
                Some(&self.nvda),
            ) == ID_DIALOG_YES;
            if remember {
                self.remember_sftp_sudo_password(resource, &trimmed);
            }
            trimmed
        };

        let mut elevated = resource.clone();
        elevated.sudo_password = password;
        Some(elevated)
    }

    unsafe fn promote_sftp_remote_for_retry(
        &mut self,
        remote: &mut Option<RemoteLocation>,
        force_prompt: bool,
    ) -> bool {
        let Some(current) = remote.clone() else {
            return false;
        };
        if current.resource.protocol != NetworkProtocol::Sftp {
            return false;
        }
        if current.resource.username.eq_ignore_ascii_case("root") && !force_prompt {
            return false;
        }
        let Some(resource) = self.resolve_sftp_root_resource(&current.resource, force_prompt)
        else {
            return false;
        };
        *remote = Some(RemoteLocation {
            resource,
            original_resource: current.original_resource.clone(),
            path: current.path,
            sftp_privilege_mode: SftpPrivilegeMode::Root,
        });
        true
    }

    unsafe fn prepare_sftp_sudo_retry(
        &mut self,
        remote: &RemoteLocation,
        force_prompt: bool,
    ) -> Option<RemoteLocation> {
        let resource = self.resolve_sftp_sudo_resource(&remote.original_resource, force_prompt)?;
        Some(RemoteLocation {
            resource,
            original_resource: remote.original_resource.clone(),
            path: remote.path.clone(),
            sftp_privilege_mode: SftpPrivilegeMode::Sudo,
        })
    }

    unsafe fn next_sftp_retry_remote(
        &mut self,
        current: &RemoteLocation,
        error_message: &str,
    ) -> Option<RemoteLocation> {
        if current.resource.protocol != NetworkProtocol::Sftp
            || !is_sftp_retry_message(error_message)
        {
            return None;
        }

        if current.uses_sftp_sudo() {
            return is_sudo_auth_failure_message(error_message)
                .then(|| self.prepare_sftp_sudo_retry(current, true))
                .flatten();
        }

        if current.uses_sftp_root() || current.resource.username.eq_ignore_ascii_case("root") {
            if is_auth_failure_message(error_message) {
                let mut root_retry = Some(current.clone());
                if self.promote_sftp_remote_for_retry(&mut root_retry, true) {
                    return root_retry;
                }
                return self.prepare_sftp_sudo_retry(current, false);
            }
            return None;
        }

        if is_auth_failure_message(error_message) {
            return None;
        }

        if !current.resource.username.trim().is_empty()
            && !current.resource.username.eq_ignore_ascii_case("root")
        {
            if let Some(sudo_retry) = self.prepare_sftp_sudo_retry(current, false) {
                return Some(sudo_retry);
            }
        }

        let mut root_retry = Some(current.clone());
        if self.promote_sftp_remote_for_retry(&mut root_retry, false) {
            return root_retry;
        }
        None
    }

    unsafe fn confirm_elevated_operation(
        &mut self,
        operation_label: &str,
        details: Vec<String>,
    ) -> bool {
        let mut lines = vec![
            "Próbujesz wykonać operację wymagającą uprawnień administratora.".to_string(),
            format!("Operacja: {operation_label}"),
        ];
        lines.extend(details.into_iter().filter(|line| !line.trim().is_empty()));
        lines.push("Czy chcesz przyznać uprawnienia tej operacji?".to_string());
        show_choice_prompt(
            self.hwnd,
            "Uprawnienia administratora",
            lines,
            vec![
                DialogButton {
                    id: ID_DIALOG_OK,
                    label: "Uzyskaj uprawnienia",
                    is_default: true,
                },
                DialogButton {
                    id: ID_DIALOG_CANCEL,
                    label: "Anuluj",
                    is_default: false,
                },
            ],
            Some(&self.nvda),
        ) == ID_DIALOG_OK
    }

    fn selected_network_resource(&self, panel_index: usize) -> Option<NetworkResource> {
        self.panels[panel_index]
            .selected_entry()
            .and_then(|entry| entry.network_resource.clone())
    }

    unsafe fn update_open_remote_resources(
        &mut self,
        old_key: &str,
        replacement: Option<&NetworkResource>,
    ) {
        for panel in &mut self.panels {
            if let PanelLocation::Remote(remote) = &mut panel.location
                && remote.resource.stable_key() == old_key
            {
                if let Some(resource) = replacement {
                    remote.resource = resource.clone();
                } else {
                    panel.location = PanelLocation::NetworkResources;
                }
            }
        }
    }

    unsafe fn edit_selected_network_resource(&mut self, panel_index: usize) {
        let Some(existing) = self.selected_network_resource(panel_index) else {
            self.notify("wybierz zapisane połączenie sieciowe");
            return;
        };
        let old_key = existing.stable_key();
        self.notify("edytuj połączenie sieciowe");
        let Some(updated) =
            show_network_connection_dialog(self.hwnd, Some(existing.clone()), Some(&self.nvda))
        else {
            self.notify("anulowano");
            return;
        };
        let new_key = updated.stable_key();
        if new_key != old_key
            && self
                .settings
                .network_resources
                .iter()
                .any(|resource| resource.stable_key() == new_key)
        {
            self.notify("połączenie sieciowe już istnieje");
            return;
        }
        if let Some(resource) = self
            .settings
            .network_resources
            .iter_mut()
            .find(|resource| resource.stable_key() == old_key)
        {
            *resource = updated.clone();
        }
        self.settings.network_resources.sort_by(|left, right| {
            left.effective_display_name()
                .to_lowercase()
                .cmp(&right.effective_display_name().to_lowercase())
        });
        self.update_open_remote_resources(&old_key, Some(&updated));
        self.save_settings_or_report();
        self.refresh_network_panels();
        self.notify(format!(
            "zapisano połączenie sieciowe {}",
            updated.effective_display_name()
        ));
    }

    unsafe fn remove_selected_network_resource(&mut self, panel_index: usize) {
        let Some(resource) = self.selected_network_resource(panel_index) else {
            self.notify("wybierz zapisane połączenie sieciowe");
            return;
        };
        let accepted = show_operation_prompt(
            self.hwnd,
            "Usuń połączenie sieciowe",
            vec![
                "Czy usunąć zapisane połączenie sieciowe?".to_string(),
                resource.summary_line(),
            ],
            Some(&self.nvda),
        );
        if !accepted {
            self.notify("anulowano");
            return;
        }
        let old_key = resource.stable_key();
        self.settings
            .network_resources
            .retain(|entry| entry.stable_key() != old_key);
        self.update_open_remote_resources(&old_key, None);
        self.save_settings_or_report();
        self.refresh_network_panels();
        self.notify(format!(
            "usunięto połączenie sieciowe {}",
            resource.effective_display_name()
        ));
    }

    unsafe fn show_network_resource_dialog(&mut self, initial: Option<NetworkResource>) {
        self.notify(if initial.is_some() {
            "edytuj połączenie sieciowe"
        } else {
            "dodaj połączenie sieciowe"
        });
        if let Some(resource) = show_network_connection_dialog(self.hwnd, initial, Some(&self.nvda))
        {
            self.add_network_resource(resource);
        } else {
            self.notify("anulowano");
        }
    }

    unsafe fn discover_network_servers(&mut self) {
        self.notify("trwa wyszukiwanie serwerów w sieci lokalnej");
        let shared_scan = Arc::new(Mutex::new(None::<DiscoveryScanResult>));
        let worker_scan = shared_scan.clone();
        let cache_snapshot = self.discovery_cache.clone();
        let initial_lines = build_discovery_progress_lines(
            "Przygotowanie skanowania.",
            None,
            0,
            0,
            0,
            0,
            Duration::ZERO,
        );
        let outcome = match run_progress_dialog(
            self.hwnd,
            "Wyszukiwanie serwerów w sieci lokalnej",
            initial_lines,
            Some(&self.nvda),
            ProgressDialogOptions::default(),
            move |progress_hwnd, sender, cancel_flag, _conflict_receiver| {
                let mut send_lines = |lines: Vec<String>| {
                    let _ = sender.send(ProgressEvent::Update(lines));
                    unsafe {
                        PostMessageW(progress_hwnd, WM_PROGRESS_EVENT, 0, 0);
                    }
                };
                match discover_local_servers_with_progress(
                    cancel_flag.clone(),
                    cache_snapshot,
                    &mut send_lines,
                ) {
                    Ok(scan) => {
                        let found = scan.servers.len();
                        if let Ok(mut shared) = worker_scan.lock() {
                            *shared = Some(scan.clone());
                        }
                        let _ = sender.send(ProgressEvent::Finished {
                            lines: build_discovery_progress_lines(
                                if found == 0 {
                                    "Nie znaleziono serwerów w sieci lokalnej."
                                } else {
                                    "Wyszukiwanie ukończone."
                                },
                                None,
                                scan.processed_hosts,
                                scan.total_hosts,
                                found,
                                scan.cache_hits,
                                scan.elapsed,
                            ),
                            speech: if found == 0 {
                                "nie znaleziono serwerów w sieci lokalnej".to_string()
                            } else {
                                format!("znaleziono {}", pluralized_elements(found))
                            },
                            outcome: WorkerOutcome::Success,
                        });
                        unsafe {
                            PostMessageW(progress_hwnd, WM_PROGRESS_EVENT, 0, 0);
                        }
                    }
                    Err(error)
                        if error.kind() == io::ErrorKind::Interrupted
                            || cancel_flag.load(Ordering::Relaxed) =>
                    {
                        let _ = sender.send(ProgressEvent::Finished {
                            lines: build_discovery_progress_lines(
                                "Wyszukiwanie anulowane.",
                                None,
                                0,
                                0,
                                0,
                                0,
                                Duration::ZERO,
                            ),
                            speech: "anulowano".to_string(),
                            outcome: WorkerOutcome::Canceled,
                        });
                        unsafe {
                            PostMessageW(progress_hwnd, WM_PROGRESS_EVENT, 0, 0);
                        }
                    }
                    Err(error) => {
                        let _ = sender.send(ProgressEvent::Finished {
                            lines: vec![
                                "Operacja: wyszukiwanie serwerów".to_string(),
                                format!("Błąd: {error}"),
                            ],
                            speech: format!("wyszukiwanie zakończone błędem: {error}"),
                            outcome: WorkerOutcome::Error(error.to_string()),
                        });
                        unsafe {
                            PostMessageW(progress_hwnd, WM_PROGRESS_EVENT, 0, 0);
                        }
                    }
                }
            },
        ) {
            Ok(outcome) => outcome,
            Err(error) => {
                self.report_error(error);
                return;
            }
        };

        match outcome {
            WorkerOutcome::Success => {
                let Some(scan) = shared_scan.lock().ok().and_then(|scan| scan.clone()) else {
                    self.notify("nie znaleziono danych skanowania");
                    return;
                };
                self.discovery_cache.update_from_scan(&scan);
                self.save_settings_or_report();
                let servers = scan.servers;
                if servers.is_empty() {
                    self.notify("nie znaleziono serwerów w sieci lokalnej");
                    show_info_prompt(
                        self.hwnd,
                        "Wyszukiwanie serwerów",
                        vec!["Nie znaleziono serwerów w sieci lokalnej.".to_string()],
                        Some(&self.nvda),
                    );
                    return;
                }
                self.notify(format!("znaleziono {}", pluralized_elements(servers.len())));
                if let Some(server) = show_discovery_dialog(self.hwnd, servers, Some(&self.nvda)) {
                    self.show_network_resource_dialog(Some(server.to_resource_template()));
                } else {
                    self.notify("anulowano");
                }
            }
            WorkerOutcome::Canceled => self.notify("anulowano"),
            WorkerOutcome::Error(message) => self.report_error(message),
        }
    }

    unsafe fn open_remote_file_in_background(
        &mut self,
        remote: RemoteLocation,
        path: PathBuf,
        entry_name: String,
    ) {
        let mut remote = remote;
        if let Some(stream_target) = remote_media_stream_target(&remote.resource, &path) {
            match open_with_system_target(&stream_target) {
                Ok(()) => {
                    self.notify_non_interrupting(format!("otwieram strumień {}", entry_name));
                    return;
                }
                Err(_) => {}
            }
        }

        loop {
            let worker_entry_name = entry_name.clone();
            let shared_path = Arc::new(Mutex::new(None::<PathBuf>));
            let worker_path = shared_path.clone();
            let remote_for_worker = remote.clone();
            let path_for_worker = path.clone();
            let initial_lines = vec![
                "Operacja: pobieranie".to_string(),
                "Przygotowanie pobierania pliku.".to_string(),
                format!("Źródło: {}", path.display()),
                "Dane: ustalanie rozmiaru".to_string(),
            ];
            let outcome = match run_progress_dialog(
                self.hwnd,
                "Pobieranie pliku z zasobu",
                initial_lines,
                Some(&self.nvda),
                ProgressDialogOptions {
                    auto_close_on_success: true,
                    auto_close_on_retryable_error: true,
                },
                move |progress_hwnd, sender, cancel_flag, conflict_receiver| {
                    let mut progress = ProgressReporter::new(
                        sender,
                        progress_hwnd,
                        cancel_flag,
                        "pobieranie",
                        ItemCounts::default(),
                        conflict_receiver,
                    );
                    match download_remote_file_to_temp_with_progress(
                        &remote_for_worker,
                        &path_for_worker,
                        &mut progress,
                    ) {
                        Ok(local_path) => {
                            if let Ok(mut shared) = worker_path.lock() {
                                *shared = Some(local_path);
                            }
                            progress.finish(
                                "Pobieranie ukończone.",
                                format!("pobrano {}", worker_entry_name),
                                WorkerOutcome::Success,
                            );
                        }
                        Err(error) if is_canceled_io_error(&error) || progress.is_canceled() => {
                            progress.finish(
                                "Pobieranie anulowane.",
                                "anulowano".to_string(),
                                WorkerOutcome::Canceled,
                            );
                        }
                        Err(error) => {
                            progress.finish(
                                &format!("Błąd: {error}"),
                                format!("pobieranie zakończone błędem: {error}"),
                                WorkerOutcome::Error(error.to_string()),
                            );
                        }
                    }
                },
            ) {
                Ok(outcome) => outcome,
                Err(error) => {
                    self.report_error(error);
                    return;
                }
            };

            match outcome {
                WorkerOutcome::Success => {
                    let local_copy = shared_path.lock().ok().and_then(|path| path.clone());
                    let Some(local_copy) = local_copy else {
                        self.report_error("nie udało się przygotować pliku tymczasowego");
                        return;
                    };
                    match open_with_system(&local_copy) {
                        Ok(()) => self.notify_non_interrupting(format!("otwieram {}", entry_name)),
                        Err(error) => self.report_error(error),
                    }
                    return;
                }
                WorkerOutcome::Canceled => {
                    self.notify("anulowano");
                    return;
                }
                WorkerOutcome::Error(message)
                    if is_sftp_retry_message(&message)
                        && remote.resource.protocol == NetworkProtocol::Sftp =>
                {
                    if let Some(candidate) = self.next_sftp_retry_remote(&remote, &message) {
                        remote = candidate;
                        continue;
                    }
                    self.notify("anulowano");
                    return;
                }
                WorkerOutcome::Error(message) => {
                    self.report_error(message);
                    return;
                }
            }
        }
    }

    unsafe fn open_archive_file_in_background(
        &mut self,
        archive: ArchiveLocation,
        inner_path: PathBuf,
        entry_name: String,
    ) {
        let temp_dir = match unique_temp_file_path("AmigaFmNativeArchive", "open") {
            Ok(path) => path,
            Err(error) => {
                self.report_error(error);
                return;
            }
        };
        let opened_path = temp_dir.join(&inner_path);
        let worker_archive = archive.archive_path.clone();
        let worker_inner = inner_path.clone();
        let worker_temp = temp_dir.clone();
        let worker_entry_name = entry_name.clone();
        let outcome = match run_progress_dialog(
            self.hwnd,
            "Otwieranie pliku z archiwum",
            vec![
                "Operacja: wypakowanie tymczasowe".to_string(),
                format!("Archiwum: {}", worker_archive.display()),
                format!("Element: {}", archive_inner_to_7z_arg(&worker_inner)),
            ],
            Some(&self.nvda),
            ProgressDialogOptions {
                auto_close_on_success: true,
                auto_close_on_retryable_error: false,
            },
            move |progress_hwnd, sender, cancel_flag, _conflict_receiver| {
                if cancel_flag.load(Ordering::Relaxed) {
                    let _ = sender.send(ProgressEvent::Finished {
                        lines: vec!["Otwieranie anulowane.".to_string()],
                        speech: "anulowano".to_string(),
                        outcome: WorkerOutcome::Canceled,
                    });
                    unsafe {
                        PostMessageW(progress_hwnd, WM_PROGRESS_EVENT, 0, 0);
                    }
                    return;
                }
                let result =
                    extract_archive_items_to_dir(&worker_archive, &[worker_inner], &worker_temp);
                match result {
                    Ok(()) => {
                        let _ = sender.send(ProgressEvent::Finished {
                            lines: vec!["Plik wypakowany tymczasowo.".to_string()],
                            speech: format!("otwieram {}", worker_entry_name),
                            outcome: WorkerOutcome::Success,
                        });
                    }
                    Err(error) => {
                        let _ = sender.send(ProgressEvent::Finished {
                            lines: vec![format!("Błąd: {error}")],
                            speech: format!("otwieranie zakończone błędem: {error}"),
                            outcome: WorkerOutcome::Error(error.to_string()),
                        });
                    }
                }
                unsafe {
                    PostMessageW(progress_hwnd, WM_PROGRESS_EVENT, 0, 0);
                }
            },
        ) {
            Ok(outcome) => outcome,
            Err(error) => {
                self.report_error(error);
                return;
            }
        };
        match outcome {
            WorkerOutcome::Success => match open_with_system(&opened_path) {
                Ok(()) => self.notify_non_interrupting(format!("otwieram {}", entry_name)),
                Err(error) => self.report_error(error),
            },
            WorkerOutcome::Canceled => self.notify("anulowano"),
            WorkerOutcome::Error(message) => self.report_error(message),
        }
    }

    unsafe fn extract_selected_archive_items(&mut self, panel_index: usize, mode: ExtractMode) {
        let destination = match self.archive_extract_destination(panel_index, mode) {
            Ok(path) => path,
            Err(message) => {
                self.notify(message);
                return;
            }
        };
        let jobs = match self.archive_extract_jobs(panel_index, mode, &destination) {
            Ok(jobs) => jobs,
            Err(message) => {
                self.notify(message);
                return;
            }
        };
        if jobs.is_empty() {
            self.notify("brak archiwów do wypakowania");
            return;
        }
        let mut counts = ItemCounts::default();
        let mut job_summaries = Vec::new();
        for job in &jobs {
            match archive_extract_counts(&job.archive_path, &job.items) {
                Ok(summary) => {
                    counts.add(summary);
                    job_summaries.push((job.clone(), summary));
                }
                Err(error) => {
                    self.report_error(error);
                    return;
                }
            }
        }
        let accepted = show_operation_prompt(
            self.hwnd,
            "Wypakowywanie",
            vec![
                format!("Archiwa: {}", pluralized_elements(jobs.len())),
                format!("Cel: {}", destination.display()),
                format!("Obiekty: {}", counts_description(counts)),
                format!("Rozmiar: {}", format_bytes(counts.bytes)),
            ],
            Some(&self.nvda),
        );
        if !accepted {
            self.notify("anulowano");
            return;
        }
        let focus_destination = destination.clone();
        let initial_lines = build_progress_lines(
            "wypakowywanie",
            "Przygotowanie operacji.",
            None,
            Some(&destination),
            ItemCounts::default(),
            counts,
            Duration::default(),
        );
        let outcome = match run_progress_dialog(
            self.hwnd,
            "Wypakowywanie",
            initial_lines,
            Some(&self.nvda),
            ProgressDialogOptions::default(),
            move |progress_hwnd, sender, cancel_flag, _conflict_receiver| {
                let mut progress = ProgressReporter::new(
                    sender,
                    progress_hwnd,
                    cancel_flag,
                    "wypakowywanie",
                    counts,
                    _conflict_receiver,
                );
                let mut processed = ItemCounts::default();
                for (job, summary) in job_summaries {
                    if progress.is_canceled() {
                        progress.finish(
                            "Wypakowywanie anulowane.",
                            "anulowano".to_string(),
                            WorkerOutcome::Canceled,
                        );
                        return;
                    }
                    progress.set_processed(
                        processed,
                        "Trwa wypakowywanie archiwum.",
                        Some(&job.archive_path),
                        Some(&job.destination),
                    );
                    if let Err(error) = extract_archive_items_to_dir_progress(
                        &job.archive_path,
                        &job.items,
                        &job.destination,
                        processed,
                        summary,
                        &mut progress,
                    ) {
                        if is_canceled_io_error(&error) {
                            progress.finish(
                                "Wypakowywanie anulowane.",
                                "anulowano".to_string(),
                                WorkerOutcome::Canceled,
                            );
                        } else {
                            progress.finish(
                                &format!("Błąd: {error}"),
                                format!("wypakowywanie zakończone błędem: {error}"),
                                WorkerOutcome::Error(error.to_string()),
                            );
                        }
                        return;
                    }
                    processed.add(summary);
                    progress.set_processed(
                        processed,
                        "Archiwum wypakowane.",
                        Some(&job.archive_path),
                        Some(&job.destination),
                    );
                }
                progress.set_processed(counts, "Wypakowywanie ukończone.", None, None);
                progress.finish(
                    "Wypakowywanie ukończone.",
                    "wypakowywanie ukończone".to_string(),
                    WorkerOutcome::Success,
                );
            },
        ) {
            Ok(outcome) => outcome,
            Err(error) => {
                self.report_error(error);
                return;
            }
        };
        match outcome {
            WorkerOutcome::Success => {
                self.notify_non_interrupting("wypakowywanie ukończone".to_string());
                if let Some(index) = (0..2).find(|index| {
                    self.panels[*index].current_dir() == Some(focus_destination.as_path())
                }) {
                    let _ = self.refresh_and_keep(index, None);
                }
            }
            WorkerOutcome::Canceled => self.notify("anulowano"),
            WorkerOutcome::Error(message) => self.report_error(message),
        }
    }

    fn archive_extract_destination(
        &self,
        panel_index: usize,
        mode: ExtractMode,
    ) -> Result<PathBuf, String> {
        match mode {
            ExtractMode::OtherPanel => self.panels[1 - panel_index]
                .current_dir_owned()
                .ok_or_else(|| "drugi panel musi wskazywać lokalny katalog".to_string()),
            ExtractMode::Here | ExtractMode::NamedFolder => {
                match &self.panels[panel_index].location {
                    PanelLocation::Filesystem(path) => Ok(path.clone()),
                    PanelLocation::Archive(archive) => archive
                        .archive_path
                        .parent()
                        .map(Path::to_path_buf)
                        .ok_or_else(|| "nie można ustalić katalogu archiwum".to_string()),
                    _ => Err("wypakowywanie działa dla lokalnych archiwów".to_string()),
                }
            }
        }
    }

    fn archive_extract_jobs(
        &self,
        panel_index: usize,
        mode: ExtractMode,
        base_destination: &Path,
    ) -> Result<Vec<ArchiveExtractJob>, String> {
        match &self.panels[panel_index].location {
            PanelLocation::Archive(archive) => {
                let targets = self.current_targets(panel_index)?;
                let destination = if matches!(mode, ExtractMode::NamedFolder) {
                    base_destination.join(archive_output_folder_name(&archive.archive_path))
                } else {
                    base_destination.to_path_buf()
                };
                Ok(vec![ArchiveExtractJob {
                    archive_path: archive.archive_path.clone(),
                    items: targets,
                    destination,
                }])
            }
            PanelLocation::Filesystem(_) => {
                let targets = self.current_targets(panel_index)?;
                let mut jobs = Vec::new();
                for target in targets {
                    if !is_archive_file_path(&target) {
                        continue;
                    }
                    let destination = if matches!(mode, ExtractMode::NamedFolder) {
                        base_destination.join(archive_output_folder_name(&target))
                    } else {
                        base_destination.to_path_buf()
                    };
                    jobs.push(ArchiveExtractJob {
                        archive_path: target,
                        items: Vec::new(),
                        destination,
                    });
                }
                Ok(jobs)
            }
            _ => Err("wypakowywanie działa dla lokalnych archiwów".to_string()),
        }
    }

    unsafe fn create_archive_from_selection(&mut self, panel_index: usize) {
        if !matches!(
            self.panels[panel_index].location,
            PanelLocation::Filesystem(_)
        ) {
            self.notify("tworzenie archiwum działa dla lokalnych elementów");
            return;
        }
        let targets = match self.current_targets(panel_index) {
            Ok(targets) => targets,
            Err(message) => {
                self.notify(message);
                return;
            }
        };
        let Some(base_dir) = self.panels[panel_index].current_dir_owned() else {
            self.notify("panel musi wskazywać lokalny katalog");
            return;
        };
        let default_name = default_archive_name(&targets);
        let Some(options) = show_archive_create_dialog(self.hwnd, &default_name, Some(&self.nvda))
        else {
            self.notify("anulowano");
            return;
        };
        let output_path = archive_output_path(&base_dir, &options);
        let mut source_counts = ItemCounts::default();
        for target in &targets {
            match summarize_path(target) {
                Ok(summary) => source_counts.add(summary),
                Err(error) => {
                    self.report_error(error);
                    return;
                }
            }
        }
        let progress_total = archive_create_progress_total(source_counts, options.format);
        let accepted = show_operation_prompt(
            self.hwnd,
            "Tworzenie archiwum",
            vec![
                format!("Elementy: {}", summarize_targets(&targets)),
                format!("Plik: {}", output_path.display()),
                format!("Typ: {}", options.format.label()),
                format!("Obiekty: {}", counts_description(source_counts)),
                format!("Rozmiar: {}", format_bytes(source_counts.bytes)),
            ],
            Some(&self.nvda),
        );
        if !accepted {
            self.notify("anulowano");
            return;
        }
        let worker_targets = targets.clone();
        let worker_options = options.clone();
        let worker_output = output_path.clone();
        let initial_lines = build_progress_lines(
            "tworzenie archiwum",
            "Przygotowanie operacji.",
            None,
            Some(&output_path),
            ItemCounts::default(),
            progress_total,
            Duration::default(),
        );
        let outcome = match run_progress_dialog(
            self.hwnd,
            "Tworzenie archiwum",
            initial_lines,
            Some(&self.nvda),
            ProgressDialogOptions::default(),
            move |progress_hwnd, sender, cancel_flag, conflict_receiver| {
                let mut progress = ProgressReporter::new(
                    sender,
                    progress_hwnd,
                    cancel_flag,
                    "tworzenie archiwum",
                    progress_total,
                    conflict_receiver,
                );
                if progress.is_canceled() {
                    progress.finish(
                        "Tworzenie archiwum anulowane.",
                        "anulowano".to_string(),
                        WorkerOutcome::Canceled,
                    );
                    return;
                }
                match create_archive_with_7z_progress(
                    &worker_output,
                    &worker_targets,
                    &worker_options,
                    source_counts,
                    &mut progress,
                ) {
                    Ok(()) => {
                        progress.set_processed(
                            progress_total,
                            "Tworzenie archiwum ukończone.",
                            None,
                            Some(&worker_output),
                        );
                        progress.finish(
                            "Tworzenie archiwum ukończone.",
                            "tworzenie archiwum ukończone".to_string(),
                            WorkerOutcome::Success,
                        );
                    }
                    Err(error) if is_canceled_io_error(&error) => {
                        progress.finish(
                            "Tworzenie archiwum anulowane.",
                            "anulowano".to_string(),
                            WorkerOutcome::Canceled,
                        );
                    }
                    Err(error) => {
                        progress.finish(
                            &format!("Błąd: {error}"),
                            format!("tworzenie archiwum zakończone błędem: {error}"),
                            WorkerOutcome::Error(error.to_string()),
                        );
                    }
                }
            },
        ) {
            Ok(outcome) => outcome,
            Err(error) => {
                self.report_error(error);
                return;
            }
        };
        match outcome {
            WorkerOutcome::Success => {
                let _ = self.refresh_and_keep(panel_index, Some(&output_path));
                self.notify_non_interrupting("tworzenie archiwum ukończone".to_string());
            }
            WorkerOutcome::Canceled => self.notify("anulowano"),
            WorkerOutcome::Error(message) => self.report_error(message),
        }
    }

    unsafe fn join_split_archive_from_selection(&mut self, panel_index: usize) {
        if !matches!(
            self.panels[panel_index].location,
            PanelLocation::Filesystem(_)
        ) {
            self.notify("łączenie części działa dla lokalnych plików");
            return;
        }
        let mut targets = match self.current_targets(panel_index) {
            Ok(targets) => targets,
            Err(message) => {
                self.notify(message);
                return;
            }
        };
        targets.sort();
        if targets.len() < 2 {
            self.notify("zaznacz co najmniej dwie części archiwum");
            return;
        }
        let Some(base_dir) = self.panels[panel_index].current_dir_owned() else {
            self.notify("panel musi wskazywać lokalny katalog");
            return;
        };
        let default_name = joined_archive_default_name(&targets[0]);
        let Some(name) = show_input_dialog(
            self.hwnd,
            "Połącz podzielone archiwum",
            "Nazwa pliku wynikowego:",
            &default_name,
            Some(&self.nvda),
        ) else {
            self.notify("anulowano");
            return;
        };
        let output = base_dir.join(name.trim());
        let worker_targets = targets.clone();
        let worker_output = output.clone();
        let outcome = run_simple_worker_dialog(
            self.hwnd,
            "Łączenie archiwum",
            vec![
                "Operacja: łączenie części".to_string(),
                format!("Plik wynikowy: {}", output.display()),
            ],
            Some(&self.nvda),
            move || join_split_files(&worker_targets, &worker_output),
            "łączenie archiwum ukończone",
        );
        match outcome {
            Ok(WorkerOutcome::Success) => {
                let _ = self.refresh_and_keep(panel_index, Some(&output));
                self.notify_non_interrupting("łączenie archiwum ukończone".to_string());
            }
            Ok(WorkerOutcome::Canceled) => self.notify("anulowano"),
            Ok(WorkerOutcome::Error(message)) => self.report_error(message),
            Err(error) => self.report_error(error),
        }
    }

    unsafe fn create_checksum_for_selection(&mut self, panel_index: usize) {
        if !matches!(
            self.panels[panel_index].location,
            PanelLocation::Filesystem(_)
        ) {
            self.notify("sumy kontrolne działają dla lokalnych plików");
            return;
        }
        let targets = match self.current_targets(panel_index) {
            Ok(targets) => targets,
            Err(message) => {
                self.notify(message);
                return;
            }
        };
        let Some(base_dir) = self.panels[panel_index].current_dir_owned() else {
            self.notify("panel musi wskazywać lokalny katalog");
            return;
        };
        let output = checksum_output_path(&base_dir, &targets);
        let worker_targets = targets.clone();
        let worker_output = output.clone();
        let outcome = run_simple_worker_dialog(
            self.hwnd,
            "Tworzenie sum kontrolnych",
            vec![
                "Operacja: SHA-256".to_string(),
                format!("Plik: {}", output.display()),
            ],
            Some(&self.nvda),
            move || create_sha256_file(&worker_targets, &worker_output),
            "Suma kontrolna została utworzona",
        );
        match outcome {
            Ok(WorkerOutcome::Success) => {
                let _ = self.refresh_and_keep(panel_index, Some(&output));
                self.notify_non_interrupting("Suma kontrolna została utworzona".to_string());
            }
            Ok(WorkerOutcome::Canceled) => self.notify("anulowano"),
            Ok(WorkerOutcome::Error(message)) => self.report_error(message),
            Err(error) => self.report_error(error),
        }
    }

    unsafe fn verify_checksum_for_selection(&mut self, panel_index: usize) {
        if !matches!(
            self.panels[panel_index].location,
            PanelLocation::Filesystem(_)
        ) {
            self.notify("sprawdzanie sum działa dla lokalnych plików");
            return;
        }
        let checksum_file = match self.current_single_target(panel_index) {
            Ok(path) => path,
            Err(message) => {
                self.notify(message);
                return;
            }
        };
        let checksum_file = match resolve_checksum_file_for_target(&checksum_file) {
            Ok(path) => path,
            Err(error) => {
                self.report_error(error);
                return;
            }
        };
        let worker_file = checksum_file.clone();
        let verify_report = Arc::new(Mutex::new(None::<ChecksumVerifyReport>));
        let worker_report = verify_report.clone();
        let outcome = run_simple_worker_dialog(
            self.hwnd,
            "Sprawdzanie sum kontrolnych",
            vec![
                "Operacja: sprawdzanie SHA-256".to_string(),
                format!("Plik: {}", checksum_file.display()),
            ],
            Some(&self.nvda),
            move || {
                let report = verify_sha256_file(&worker_file)?;
                if let Ok(mut slot) = worker_report.lock() {
                    *slot = Some(report);
                }
                Ok(())
            },
            "Suma kontrolna jest poprawna",
        );
        match outcome {
            Ok(WorkerOutcome::Success) => {
                let message = verify_report
                    .lock()
                    .ok()
                    .and_then(|slot| slot.clone())
                    .map(|report| report.success_message())
                    .unwrap_or_else(|| "Suma kontrolna jest poprawna".to_string());
                self.notify_non_interrupting(message);
            }
            Ok(WorkerOutcome::Canceled) => self.notify("anulowano"),
            Ok(WorkerOutcome::Error(message)) => self.report_error(message),
            Err(error) => self.report_error(error),
        }
    }

    unsafe fn show_properties_for_selected(&mut self, panel_index: usize) {
        if self.remote_location_for_panel(panel_index).is_some() {
            self.notify("właściwości są dostępne tylko dla lokalnych elementów");
            return;
        }
        match self.current_single_target(panel_index) {
            Ok(path) => match open_properties_sheet(self.hwnd, &path, None) {
                Ok(()) => self.notify(format!("właściwości {}", display_name(&path))),
                Err(error) => self.report_error(error),
            },
            Err(message) => self.notify(message),
        }
    }

    unsafe fn show_permissions_for_selected(&mut self, panel_index: usize) {
        if self.remote_location_for_panel(panel_index).is_some() {
            self.notify("uprawnienia są dostępne tylko dla lokalnych elementów");
            return;
        }
        match self.current_single_target(panel_index) {
            Ok(path) => match open_properties_sheet(self.hwnd, &path, Some("Security")) {
                Ok(()) => self.notify(format!("uprawnienia {}", display_name(&path))),
                Err(error) => self.report_error(error),
            },
            Err(message) => self.notify(message),
        }
    }

    unsafe fn show_shell_context_for_selected(&mut self, panel_index: usize) {
        if self.remote_location_for_panel(panel_index).is_some() {
            self.notify("menu kontekstowe Windows działa tylko dla lokalnych elementów");
            return;
        }
        match self.current_targets(panel_index) {
            Ok(targets) => match show_shell_context_menu(
                self.hwnd,
                self.panels[panel_index].list_hwnd,
                &targets,
            ) {
                Ok(()) => {
                    let _ = self.refresh_all();
                    self.focus_panel(panel_index, false);
                    self.notify("menu kontekstowe Windows");
                }
                Err(error) => self.report_error(error),
            },
            Err(message) => self.notify(message),
        }
    }

    unsafe fn rename_selected(&mut self, panel_index: usize) {
        if let Some(remote) = self.remote_location_for_panel(panel_index) {
            self.rename_selected_remote(panel_index, &remote);
            return;
        }
        let Some(entry) = self.panels[panel_index].selected_entry().cloned() else {
            self.notify("brak elementów");
            return;
        };
        if !entry.is_operable() {
            self.notify("zmiana nazwy wymaga pliku lub katalogu");
            return;
        }
        let Some(target) = entry.path else {
            return;
        };

        let Some(new_name) = ({
            self.notify("zmień nazwę, pole edycji");
            show_input_dialog(
                self.hwnd,
                "Zmień nazwę",
                "Nowa nazwa:",
                &entry.name,
                Some(&self.nvda),
            )
        }) else {
            self.notify("anulowano");
            return;
        };
        let trimmed = new_name.trim();
        if trimmed.is_empty() {
            self.notify("nazwa nie może być pusta");
            return;
        }

        let Some(parent) = target.parent() else {
            self.notify("nie można zmienić nazwy");
            return;
        };
        let destination = parent.join(trimmed);
        match fs::rename(&target, &destination) {
            Ok(()) => {
                if self.panels[panel_index].marked.remove(&target) {
                    self.panels[panel_index].marked.insert(destination.clone());
                }
                if let Err(error) = self.refresh_and_keep(panel_index, Some(destination.as_path()))
                {
                    self.report_error(error);
                } else {
                    self.notify_non_interrupting(format!(
                        "zmieniono nazwę na {}",
                        display_name(&destination)
                    ));
                }
            }
            Err(error) if is_permission_denied_error(&error) => {
                if !self.confirm_elevated_operation(
                    "zmiana nazwy",
                    vec![
                        format!("Element: {}", display_name(&target)),
                        format!("Nowa nazwa: {}", display_name(&destination)),
                    ],
                ) {
                    self.notify("anulowano");
                    return;
                }
                match run_elevated_local_operation(ElevatedLocalOperation::Rename {
                    source: target.clone(),
                    destination: destination.clone(),
                }) {
                    Ok(()) => {
                        if self.panels[panel_index].marked.remove(&target) {
                            self.panels[panel_index].marked.insert(destination.clone());
                        }
                        if let Err(error) =
                            self.refresh_and_keep(panel_index, Some(destination.as_path()))
                        {
                            self.report_error(error);
                        } else {
                            self.notify_non_interrupting(format!(
                                "zmieniono nazwę na {}",
                                display_name(&destination)
                            ));
                        }
                    }
                    Err(error) if is_uac_canceled_error(&error) => self.notify("anulowano"),
                    Err(error) => self.report_error(error),
                }
            }
            Err(error) => self.report_error(error),
        }
    }

    unsafe fn rename_selected_remote(&mut self, panel_index: usize, remote: &RemoteLocation) {
        let Some(entry) = self.panels[panel_index].selected_entry().cloned() else {
            self.notify("brak elementów");
            return;
        };
        if !entry.is_operable() {
            self.notify("zmiana nazwy wymaga pliku lub katalogu");
            return;
        }
        let Some(target) = entry.path else {
            self.notify("nie można zmienić nazwy");
            return;
        };

        let Some(new_name) = ({
            self.notify("zmień nazwę, pole edycji");
            show_input_dialog(
                self.hwnd,
                "Zmień nazwę",
                "Nowa nazwa:",
                &entry.name,
                Some(&self.nvda),
            )
        }) else {
            self.notify("anulowano");
            return;
        };
        let trimmed = new_name.trim();
        if trimmed.is_empty() {
            self.notify("nazwa nie może być pusta");
            return;
        }
        let Some(parent) = target.parent() else {
            self.notify("nie można zmienić nazwy");
            return;
        };
        let destination = parent.join(trimmed);
        let mut remote = Some(remote.clone());
        let result = loop {
            let current = remote.clone().expect("remote location available");
            match remote_rename(&current, &target, &destination) {
                Ok(()) => break Ok(()),
                Err(error) if is_sftp_retry_error(&error) => {
                    if let Some(candidate) =
                        self.next_sftp_retry_remote(&current, &error.to_string())
                    {
                        remote = Some(candidate);
                        continue;
                    }
                    break Err(error);
                }
                Err(error) => break Err(error),
            }
        };
        match result {
            Ok(()) => {
                if self.panels[panel_index].marked.remove(&target) {
                    self.panels[panel_index].marked.insert(destination.clone());
                }
                if let Err(error) = self.refresh_and_keep(panel_index, Some(destination.as_path()))
                {
                    self.report_error(error);
                } else {
                    self.notify_non_interrupting(format!("zmieniono nazwę na {}", trimmed));
                }
            }
            Err(error) => self.report_error(error),
        }
    }

    unsafe fn create_folder(&mut self, panel_index: usize) {
        if let Some(remote) = self.remote_location_for_panel(panel_index) {
            self.create_folder_remote(panel_index, &remote);
            return;
        }
        if matches!(self.panels[panel_index].location, PanelLocation::Archive(_)) {
            self.notify("tworzenie katalogu wewnątrz archiwum będzie dodane później");
            return;
        }
        let Some(base_dir) = self.panels[panel_index].current_dir_owned() else {
            self.notify("nowy katalog można utworzyć tylko w widoku katalogu");
            return;
        };

        let Some(name) = ({
            self.notify("nowy katalog, pole edycji");
            show_input_dialog(
                self.hwnd,
                "Nowy katalog",
                "Nazwa katalogu:",
                "",
                Some(&self.nvda),
            )
        }) else {
            self.notify("anulowano");
            return;
        };
        let trimmed = name.trim();
        if trimmed.is_empty() {
            self.notify("nazwa katalogu nie może być pusta");
            return;
        }

        let path = base_dir.join(trimmed);
        match fs::create_dir(&path) {
            Ok(()) => {
                if let Err(error) = self.refresh_and_keep(panel_index, Some(path.as_path())) {
                    self.report_error(error);
                } else {
                    self.notify_non_interrupting(format!(
                        "utworzono katalog {}",
                        display_name(&path)
                    ));
                }
            }
            Err(error) if is_permission_denied_error(&error) => {
                if !self.confirm_elevated_operation(
                    "tworzenie katalogu",
                    vec![format!("Katalog: {}", display_name(&path))],
                ) {
                    self.notify("anulowano");
                    return;
                }
                match run_elevated_local_operation(ElevatedLocalOperation::CreateDir {
                    path: path.clone(),
                }) {
                    Ok(()) => {
                        if let Err(error) = self.refresh_and_keep(panel_index, Some(path.as_path()))
                        {
                            self.report_error(error);
                        } else {
                            self.notify_non_interrupting(format!(
                                "utworzono katalog {}",
                                display_name(&path)
                            ));
                        }
                    }
                    Err(error) if is_uac_canceled_error(&error) => self.notify("anulowano"),
                    Err(error) => self.report_error(error),
                }
            }
            Err(error) => self.report_error(error),
        }
    }

    unsafe fn create_folder_remote(&mut self, panel_index: usize, remote: &RemoteLocation) {
        let Some(base_dir) = self.panels[panel_index].current_remote_dir_owned() else {
            self.notify("nowy katalog można utworzyć tylko w katalogu");
            return;
        };

        let Some(name) = ({
            self.notify("nowy katalog, pole edycji");
            show_input_dialog(
                self.hwnd,
                "Nowy katalog",
                "Nazwa katalogu:",
                "",
                Some(&self.nvda),
            )
        }) else {
            self.notify("anulowano");
            return;
        };
        let trimmed = name.trim();
        if trimmed.is_empty() {
            self.notify("nazwa katalogu nie może być pusta");
            return;
        }

        let path = base_dir.join(trimmed);
        let mut remote = Some(remote.clone());
        let result = loop {
            let current = remote.clone().expect("remote location available");
            match remote_create_dir(&current, &path) {
                Ok(()) => break Ok(()),
                Err(error) if is_sftp_retry_error(&error) => {
                    if let Some(candidate) =
                        self.next_sftp_retry_remote(&current, &error.to_string())
                    {
                        remote = Some(candidate);
                        continue;
                    }
                    break Err(error);
                }
                Err(error) => break Err(error),
            }
        };
        match result {
            Ok(()) => {
                if let Err(error) = self.refresh_and_keep(panel_index, Some(path.as_path())) {
                    self.report_error(error);
                } else {
                    self.notify_non_interrupting(format!("utworzono katalog {}", trimmed));
                }
            }
            Err(error) => self.report_error(error),
        }
    }

    unsafe fn delete_selected(&mut self, panel_index: usize) {
        if let Some(remote) = self.remote_location_for_panel(panel_index) {
            self.delete_selected_remote(panel_index, &remote);
            return;
        }
        if let PanelLocation::Archive(archive) = self.panels[panel_index].location.clone() {
            self.delete_selected_archive(panel_index, archive);
            return;
        }
        let targets = match self.current_targets(panel_index) {
            Ok(targets) => targets,
            Err(message) => {
                self.notify(message);
                return;
            }
        };

        let mut counts = ItemCounts::default();
        let mut target_summaries = Vec::new();
        for target in &targets {
            match summarize_path(target) {
                Ok(summary) => {
                    counts.add(summary);
                    target_summaries.push((target.clone(), summary));
                }
                Err(error) => {
                    self.report_error(error);
                    return;
                }
            }
        }

        let accepted = show_operation_prompt(
            self.hwnd,
            "Usuwanie",
            vec![
                "Operacja: usuwanie".to_string(),
                format!("Elementy: {}", summarize_targets(&targets)),
                format!("Obiekty: {}", counts_description(counts)),
                format!("Rozmiar: {}", format_bytes(counts.bytes)),
            ],
            Some(&self.nvda),
        );
        if !accepted {
            self.notify("anulowano");
            return;
        }

        let initial_lines = build_progress_lines(
            "usuwanie",
            "Przygotowanie operacji.",
            None,
            None,
            ItemCounts::default(),
            counts,
            Duration::default(),
        );
        let progress_targets = target_summaries.clone();
        let outcome = match run_progress_dialog(
            self.hwnd,
            "Postęp usuwania",
            initial_lines,
            Some(&self.nvda),
            ProgressDialogOptions::default(),
            move |progress_hwnd, sender, cancel_flag, conflict_receiver| {
                let mut progress = ProgressReporter::new(
                    sender,
                    progress_hwnd,
                    cancel_flag,
                    "usuwanie",
                    counts,
                    conflict_receiver,
                );
                for (target, _) in &progress_targets {
                    match delete_path_with_progress(target, &mut progress) {
                        Ok(OperationResult::Done | OperationResult::Skipped) => {}
                        Ok(OperationResult::Canceled) => {
                            progress.finish(
                                "Usuwanie anulowane.",
                                "usuwanie anulowane".to_string(),
                                WorkerOutcome::Canceled,
                            );
                            return;
                        }
                        Err(error) => {
                            progress.finish(
                                &format!("Błąd: {error}"),
                                format!("operacja zakończona błędem: {error}"),
                                WorkerOutcome::Error(error.to_string()),
                            );
                            return;
                        }
                    }
                }
                progress.finish(
                    "Usuwanie ukończone.",
                    "usuwanie ukończone".to_string(),
                    WorkerOutcome::Success,
                );
            },
        ) {
            Ok(outcome) => outcome,
            Err(error) => {
                self.report_error(error);
                return;
            }
        };

        match outcome {
            WorkerOutcome::Success => {}
            WorkerOutcome::Canceled => {
                self.notify("anulowano");
                return;
            }
            WorkerOutcome::Error(message) if is_permission_denied_message(&message) => {
                if !self.confirm_elevated_operation(
                    "usuwanie",
                    vec![
                        format!("Elementy: {}", summarize_targets(&targets)),
                        format!("Obiekty: {}", counts_description(counts)),
                    ],
                ) {
                    self.notify("anulowano");
                    return;
                }
                match run_elevated_local_operation(ElevatedLocalOperation::DeleteTargets {
                    targets: targets.clone(),
                }) {
                    Ok(()) => {}
                    Err(error) if is_uac_canceled_error(&error) => {
                        self.notify("anulowano");
                        return;
                    }
                    Err(error) => {
                        self.report_error(error);
                        return;
                    }
                }
            }
            WorkerOutcome::Error(message) => {
                self.report_error(message);
                return;
            }
        }

        for target in &targets {
            if let Err(error) = validate_path_removed(target) {
                self.report_error(error);
                return;
            }
        }

        self.panels[panel_index]
            .marked
            .retain(|path| !targets.iter().any(|target| path == target));
        if let Err(error) = self.refresh_and_keep(panel_index, None) {
            self.report_error(error);
            return;
        }

        self.notify_non_interrupting(format!(
            "usuwanie ukończone, {}",
            counts_description(counts)
        ));
    }

    unsafe fn delete_selected_archive(&mut self, panel_index: usize, archive: ArchiveLocation) {
        let targets = match self.current_targets(panel_index) {
            Ok(targets) => targets,
            Err(message) => {
                self.notify(message);
                return;
            }
        };
        let accepted = show_operation_prompt(
            self.hwnd,
            "Usuwanie z archiwum",
            vec![
                "Operacja: usuwanie z archiwum".to_string(),
                format!("Elementy: {}", summarize_targets(&targets)),
                format!("Archiwum: {}", archive.archive_path.display()),
            ],
            Some(&self.nvda),
        );
        if !accepted {
            self.notify("anulowano");
            return;
        }
        let worker_archive = archive.archive_path.clone();
        let worker_targets = targets.clone();
        let outcome = run_simple_worker_dialog(
            self.hwnd,
            "Usuwanie z archiwum",
            vec![
                "Operacja: usuwanie z archiwum".to_string(),
                format!("Archiwum: {}", archive.archive_path.display()),
            ],
            Some(&self.nvda),
            move || delete_archive_items(&worker_archive, &worker_targets),
            "usuwanie z archiwum ukończone",
        );
        match outcome {
            Ok(WorkerOutcome::Success) => {
                self.panels[panel_index]
                    .marked
                    .retain(|path| !targets.iter().any(|target| path == target));
                let _ = self.refresh_and_keep(panel_index, None);
                self.notify_non_interrupting("usuwanie z archiwum ukończone".to_string());
            }
            Ok(WorkerOutcome::Canceled) => self.notify("anulowano"),
            Ok(WorkerOutcome::Error(message)) => self.report_error(message),
            Err(error) => self.report_error(error),
        }
    }

    unsafe fn delete_selected_remote(&mut self, panel_index: usize, remote: &RemoteLocation) {
        let targets = match self.current_targets(panel_index) {
            Ok(targets) => targets,
            Err(message) => {
                self.notify(message);
                return;
            }
        };

        let mut counts = ItemCounts::default();
        for target in &targets {
            if let Some(entry) = self.panels[panel_index]
                .entries
                .iter()
                .find(|entry| entry.path.as_deref() == Some(target.as_path()))
            {
                match entry.kind {
                    EntryKind::Directory => counts.directories += 1,
                    EntryKind::File => counts.files += 1,
                    _ => {}
                }
            }
        }

        let accepted = show_operation_prompt(
            self.hwnd,
            "Usuwanie",
            vec![
                "Operacja: usuwanie".to_string(),
                format!("Elementy: {}", summarize_targets(&targets)),
                format!("Obiekty: {}", counts_description(counts)),
            ],
            Some(&self.nvda),
        );
        if !accepted {
            self.notify("anulowano");
            return;
        }

        let mut sftp_fallback: Option<RemoteLocation> = None;
        let result = loop {
            match remote_delete_targets_with_sftp_fallback(
                remote,
                sftp_fallback.as_ref(),
                &targets,
                &self.panels[panel_index].entries,
            ) {
                Ok(()) => break Ok(()),
                Err(error) if is_sftp_retry_error(&error) => {
                    let candidate = if let Some(fallback) = sftp_fallback.clone() {
                        self.next_sftp_retry_remote(&fallback, &error.to_string())
                    } else {
                        self.next_sftp_retry_remote(remote, &error.to_string())
                    };
                    if let Some(candidate) = candidate {
                        sftp_fallback = Some(candidate);
                        continue;
                    }
                    break Err(error);
                }
                Err(error) => break Err(error),
            }
        };
        match result {
            Ok(()) => {
                self.panels[panel_index]
                    .marked
                    .retain(|path| !targets.iter().any(|target| path == target));
                if let Err(error) = self.refresh_and_keep(panel_index, None) {
                    self.report_error(error);
                } else {
                    self.notify_non_interrupting(format!(
                        "usuwanie ukończone, {}",
                        counts_description(counts)
                    ));
                }
            }
            Err(error) => self.report_error(error),
        }
    }

    unsafe fn copy_or_move_selected(&mut self, panel_index: usize, move_mode: bool) {
        if matches!(self.panels[panel_index].location, PanelLocation::Archive(_)) {
            if move_mode {
                self.notify("przenoszenie z archiwum nie jest jeszcze obsługiwane");
            } else {
                self.extract_selected_archive_items(panel_index, ExtractMode::OtherPanel);
            }
            return;
        }
        if let PanelLocation::Archive(archive) = self.panels[1 - panel_index].location.clone()
            && matches!(
                self.panels[panel_index].location,
                PanelLocation::Filesystem(_)
            )
        {
            self.add_selected_to_archive(panel_index, 1 - panel_index, archive, move_mode);
            return;
        }
        let targets = match self.current_targets(panel_index) {
            Ok(targets) => targets,
            Err(message) => {
                self.notify(message);
                return;
            }
        };
        let passive_index = 1 - panel_index;
        self.transfer_targets(
            Some(panel_index),
            passive_index,
            self.remote_location_for_panel(panel_index),
            self.panels[panel_index].entries.clone(),
            targets,
            move_mode,
            true,
        );
    }

    unsafe fn add_selected_to_archive(
        &mut self,
        source_panel_index: usize,
        archive_panel_index: usize,
        archive: ArchiveLocation,
        move_mode: bool,
    ) {
        let targets = match self.current_targets(source_panel_index) {
            Ok(targets) => targets,
            Err(message) => {
                self.notify(message);
                return;
            }
        };
        let Some(base_dir) = self.panels[source_panel_index].current_dir_owned() else {
            self.notify("źródłem musi być lokalny katalog");
            return;
        };
        let accepted = show_operation_prompt(
            self.hwnd,
            "Dodawanie do archiwum",
            vec![
                format!("Elementy: {}", summarize_targets(&targets)),
                format!("Archiwum: {}", archive.archive_path.display()),
            ],
            Some(&self.nvda),
        );
        if !accepted {
            self.notify("anulowano");
            return;
        }
        let worker_targets = targets.clone();
        let worker_archive = archive.archive_path.clone();
        let worker_base = base_dir.clone();
        let outcome = run_simple_worker_dialog(
            self.hwnd,
            "Dodawanie do archiwum",
            vec![
                "Operacja: dodawanie do archiwum".to_string(),
                format!("Archiwum: {}", archive.archive_path.display()),
            ],
            Some(&self.nvda),
            move || {
                add_files_to_archive(&worker_archive, &worker_targets, &worker_base)?;
                if move_mode {
                    for target in &worker_targets {
                        delete_path_after_move(target)?;
                    }
                }
                Ok(())
            },
            "dodawanie do archiwum ukończone",
        );
        match outcome {
            Ok(WorkerOutcome::Success) => {
                let _ = self.refresh_and_keep(archive_panel_index, None);
                if move_mode {
                    let _ = self.refresh_and_keep(source_panel_index, None);
                }
                self.notify_non_interrupting("dodawanie do archiwum ukończone".to_string());
            }
            Ok(WorkerOutcome::Canceled) => self.notify("anulowano"),
            Ok(WorkerOutcome::Error(message)) => self.report_error(message),
            Err(error) => self.report_error(error),
        }
    }

    unsafe fn handle_command(&mut self, wparam: WPARAM, lparam: LPARAM) {
        let code = hiword(wparam as u32);
        let id = loword(wparam as u32) as i32;

        if lparam != 0 {
            if id == IDC_LEFT_LIST || id == IDC_RIGHT_LIST {
                let panel_index = if id == IDC_LEFT_LIST { 0 } else { 1 };
                match code {
                    x if x == LBN_SETFOCUS as u16 => {
                        self.activate_panel(panel_index, true);
                    }
                    x if x == LBN_SELCHANGE as u16 => {
                        self.sync_selection_from_control(panel_index);
                        self.activate_panel(panel_index, false);
                        self.announce_selection();
                    }
                    x if x == LBN_DBLCLK as u16 => {
                        self.sync_selection_from_control(panel_index);
                        self.handle_panel_action(panel_index, PanelAction::Open);
                    }
                    _ => {}
                }
            }
            return;
        }

        match id as u16 {
            IDM_RENAME => self.handle_panel_action(self.active_panel, PanelAction::Rename),
            IDM_COPY => self.handle_panel_action(self.active_panel, PanelAction::Copy),
            IDM_MOVE => self.handle_panel_action(self.active_panel, PanelAction::Move),
            IDM_NEW_FOLDER => self.handle_panel_action(self.active_panel, PanelAction::NewFolder),
            IDM_DELETE => self.handle_panel_action(self.active_panel, PanelAction::Delete),
            IDM_MARK_ALL => self.handle_panel_action(self.active_panel, PanelAction::MarkAll),
            IDM_UNMARK_ALL => self.handle_panel_action(self.active_panel, PanelAction::UnmarkAll),
            IDM_INVERT_MARKS => {
                self.handle_panel_action(self.active_panel, PanelAction::InvertMarks)
            }
            IDM_MARK_EXTENSION => {
                self.handle_panel_action(self.active_panel, PanelAction::MarkByExtension)
            }
            IDM_MARK_NAME => self.handle_panel_action(self.active_panel, PanelAction::MarkByName),
            IDM_ADD_TO_FAVORITES => self.add_to_favorites(self.active_panel),
            IDM_ADD_NETWORK_CONNECTION => self.show_network_resource_dialog(None),
            IDM_DISCOVER_NETWORK_SERVERS => self.discover_network_servers(),
            IDM_EDIT_NETWORK_CONNECTION => self.edit_selected_network_resource(self.active_panel),
            IDM_REMOVE_NETWORK_CONNECTION => {
                self.remove_selected_network_resource(self.active_panel)
            }
            IDM_PROPERTIES => self.show_properties_for_selected(self.active_panel),
            IDM_PERMISSIONS => self.show_permissions_for_selected(self.active_panel),
            IDM_SYSTEM_CONTEXT => self.show_shell_context_for_selected(self.active_panel),
            IDM_EXTRACT_HERE => {
                self.extract_selected_archive_items(self.active_panel, ExtractMode::Here)
            }
            IDM_EXTRACT_TO_FOLDER => {
                self.extract_selected_archive_items(self.active_panel, ExtractMode::NamedFolder)
            }
            IDM_EXTRACT_TO_OTHER_PANEL => {
                self.extract_selected_archive_items(self.active_panel, ExtractMode::OtherPanel)
            }
            IDM_CREATE_ARCHIVE => self.create_archive_from_selection(self.active_panel),
            IDM_JOIN_SPLIT_ARCHIVE => self.join_split_archive_from_selection(self.active_panel),
            IDM_CHECKSUM_CREATE => self.create_checksum_for_selection(self.active_panel),
            IDM_CHECKSUM_VERIFY => self.verify_checksum_for_selection(self.active_panel),
            IDM_SEARCH => self.handle_panel_action(self.active_panel, PanelAction::Search),
            IDM_REFRESH => self.handle_panel_action(self.active_panel, PanelAction::Refresh),
            IDM_VIEW_SIZE => {
                self.view_options.show_size = !self.view_options.show_size;
                self.save_settings_or_report();
                if let Err(error) = self.refresh_view_options() {
                    self.report_error(error);
                }
            }
            IDM_VIEW_TYPE => {
                self.view_options.show_type = !self.view_options.show_type;
                self.save_settings_or_report();
                if let Err(error) = self.refresh_view_options() {
                    self.report_error(error);
                }
            }
            IDM_VIEW_CREATED => {
                self.view_options.show_created = !self.view_options.show_created;
                self.save_settings_or_report();
                if let Err(error) = self.refresh_view_options() {
                    self.report_error(error);
                }
            }
            IDM_VIEW_MODIFIED => {
                self.view_options.show_modified = !self.view_options.show_modified;
                self.save_settings_or_report();
                if let Err(error) = self.refresh_view_options() {
                    self.report_error(error);
                }
            }
            IDM_QUIT => {
                self.notify("czy chcesz wyjść z Amiga FM?");
                if confirm_exit(self.hwnd) {
                    DestroyWindow(self.hwnd);
                }
            }
            IDM_OPTIONS => {
                show_info_prompt(
                    self.hwnd,
                    "Opcje",
                    vec!["Opcje będą dodane w kolejnym kroku.".to_string()],
                    Some(&self.nvda),
                );
                self.notify("opcje będą dodane później");
            }
            IDM_ABOUT => {
                show_info_prompt(
                    self.hwnd,
                    "O programie",
                    vec![
                        "Amiga FM".to_string(),
                        "Lekki menadżer plików Win32 z obsługą NVDA napisany w rust.".to_string(),
                    ],
                    Some(&self.nvda),
                );
                self.notify("o programie");
            }
            _ => {}
        }
    }

    unsafe fn cleanup(&mut self) {
        if !self.font.is_null() {
            DeleteObject(self.font as _);
            self.font = null_mut();
        }
        if !self.background_brush.is_null() {
            DeleteObject(self.background_brush as _);
            self.background_brush = null_mut();
        }
    }
}

fn main() {
    if let Some(code) = maybe_handle_elevated_local_request() {
        std::process::exit(code);
    }
    if let Err(error) = run() {
        unsafe {
            MessageBoxW(
                null_mut(),
                wide(&error).as_ptr(),
                wide("Amiga FM").as_ptr(),
                MB_OK | MB_ICONERROR,
            );
        }
    }
}

fn run() -> Result<(), String> {
    unsafe {
        let app_id = wide("pl.turek.AmigaFM");
        let app_id_result = SetCurrentProcessExplicitAppUserModelID(app_id.as_ptr());
        if app_id_result < 0 {
            return Err(format!(
                "nie udało się ustawić identyfikatora aplikacji: 0x{app_id_result:08X}"
            ));
        }

        let mut common_controls = INITCOMMONCONTROLSEX {
            dwSize: std::mem::size_of::<INITCOMMONCONTROLSEX>() as u32,
            dwICC: ICC_STANDARD_CLASSES,
        };
        InitCommonControlsEx(&mut common_controls);

        register_class(MAIN_CLASS, Some(main_window_proc))?;
        register_class(INPUT_DIALOG_CLASS, Some(input_dialog_proc))?;
        register_class(SEARCH_DIALOG_CLASS, Some(search_dialog_proc))?;
        register_class(NETWORK_DIALOG_CLASS, Some(network_dialog_proc))?;
        register_class(DISCOVERY_DIALOG_CLASS, Some(discovery_dialog_proc))?;
        register_class(
            OPERATION_PROMPT_DIALOG_CLASS,
            Some(operation_prompt_dialog_proc),
        )?;
        register_class(PROGRESS_DIALOG_CLASS, Some(progress_dialog_proc))?;
        register_class(
            ARCHIVE_CREATE_DIALOG_CLASS,
            Some(archive_create_dialog_proc),
        )?;

        let app = Box::new(AppState::new().map_err(|error| error.to_string())?);
        let app_ptr = Box::into_raw(app);
        let instance = GetModuleHandleW(null());

        let hwnd = CreateWindowExW(
            0,
            wide(MAIN_CLASS).as_ptr(),
            wide("Amiga FM").as_ptr(),
            WS_OVERLAPPEDWINDOW | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            1100,
            700,
            null_mut(),
            null_mut(),
            instance,
            app_ptr as _,
        );

        if hwnd.is_null() {
            let _ = Box::from_raw(app_ptr);
            return Err(last_error_message("nie udało się utworzyć okna"));
        }

        SetWindowTextW(hwnd, wide("Amiga FM").as_ptr());
        ShowWindow(hwnd, SW_MAXIMIZE);
        UpdateWindow(hwnd);

        let mut message: MSG = std::mem::zeroed();
        while GetMessageW(&mut message, null_mut(), 0, 0) > 0 {
            if handle_main_clipboard_shortcut(hwnd, &message) {
                continue;
            }
            TranslateMessage(&message);
            DispatchMessageW(&message);
        }
    }

    Ok(())
}

unsafe fn handle_main_clipboard_shortcut(main_hwnd: HWND, message: &MSG) -> bool {
    let action = match message.message {
        WM_KEYDOWN if ctrl_pressed() => match message.wParam as u32 {
            0x43 => Some(PanelAction::ClipboardCopy),
            0x58 => Some(PanelAction::ClipboardCut),
            0x56 => Some(PanelAction::ClipboardPaste),
            _ => None,
        },
        WM_COPY_MSG => Some(PanelAction::ClipboardCopy),
        WM_CUT_MSG => Some(PanelAction::ClipboardCut),
        WM_PASTE_MSG => Some(PanelAction::ClipboardPaste),
        _ => None,
    };
    let Some(action) = action else {
        return false;
    };

    let Some(app) = app_state_mut(main_hwnd) else {
        return false;
    };
    let focused = GetFocus();
    let panel_index = (0..2)
        .find(|index| {
            focused == app.panels[*index].list_hwnd || message.hwnd == app.panels[*index].list_hwnd
        })
        .unwrap_or(app.active_panel);

    app.handle_panel_action(panel_index, action);
    true
}

unsafe extern "system" fn main_window_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_NCCREATE => {
            let create =
                lparam as *const windows_sys::Win32::UI::WindowsAndMessaging::CREATESTRUCTW;
            let ptr = (*create).lpCreateParams as *mut AppState;
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, ptr as isize);
            1
        }
        WM_CREATE => {
            if let Some(app) = app_state_mut(hwnd) {
                match app.initialize_window(hwnd) {
                    Ok(()) => 0,
                    Err(error) => {
                        MessageBoxW(
                            hwnd,
                            wide(&error).as_ptr(),
                            wide("Błąd uruchamiania").as_ptr(),
                            MB_OK | MB_ICONERROR,
                        );
                        -1
                    }
                }
            } else {
                -1
            }
        }
        WM_SIZE => {
            if let Some(app) = app_state_mut(hwnd) {
                app.layout();
            }
            0
        }
        WM_SETFOCUS => {
            if let Some(app) = app_state_mut(hwnd) {
                SetFocus(app.panels[app.active_panel].list_hwnd);
            }
            0
        }
        WM_COMMAND => {
            if let Some(app) = app_state_mut(hwnd) {
                app.handle_command(wparam, lparam);
            }
            0
        }
        WM_PANEL_ACTION => {
            if let Some(action) = PanelAction::from_wparam(wparam) {
                if let Some(app) = app_state_mut(hwnd) {
                    app.handle_panel_action(lparam as usize, action);
                }
            }
            0
        }
        WM_PANEL_LOAD_EVENT => {
            if let Some(app) = app_state_mut(hwnd) {
                app.process_panel_load_messages();
            }
            0
        }
        WM_SEARCH_EVENT => {
            if let Some(app) = app_state_mut(hwnd) {
                app.process_search_messages();
            }
            0
        }
        WM_PANEL_SEARCH => {
            if let Some(app) = app_state_mut(hwnd) {
                if let Some(ch) = char::from_u32(wparam as u32) {
                    app.search_next(lparam as usize, ch);
                }
            }
            0
        }
        WM_CTLCOLORSTATIC | WM_CTLCOLOREDIT | WM_CTLCOLORLISTBOX => {
            if let Some(app) = app_state_mut(hwnd) {
                let hdc = wparam as HDC;
                SetTextColor(hdc, YELLOW);
                SetBkColor(hdc, BLACK);
                if msg == WM_CTLCOLORSTATIC {
                    SetBkMode(hdc, TRANSPARENT as i32);
                }
                return app.background_brush as LRESULT;
            }
            0
        }
        WM_CLOSE => {
            if let Some(app) = app_state_mut(hwnd) {
                app.notify("czy zamknąć program");
            }
            if confirm_exit(hwnd) {
                DestroyWindow(hwnd);
            }
            0
        }
        WM_DESTROY => {
            PostQuitMessage(0);
            0
        }
        WM_NCDESTROY => {
            let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut AppState;
            if !ptr.is_null() {
                let mut app = Box::from_raw(ptr);
                app.cleanup();
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

unsafe extern "system" fn listbox_subclass_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
    subclass_id: usize,
    ref_data: usize,
) -> LRESULT {
    let panel_index = ref_data;
    let parent = windows_sys::Win32::UI::WindowsAndMessaging::GetParent(hwnd);

    match msg {
        WM_GETDLGCODE => {
            let base = DefSubclassProc(hwnd, msg, wparam, lparam);
            return base | 0x0001 | 0x0002 | 0x0004 | 0x0080;
        }
        WM_KEYDOWN => {
            let handled = match wparam as u32 {
                0x09 => Some(PanelAction::SwitchPanel),
                0x0D => Some(PanelAction::Open),
                0x08 => Some(PanelAction::Up),
                0x1B if app_state_mut(parent)
                    .map(|app| app.panels[panel_index].is_search_active())
                    .unwrap_or(false) =>
                {
                    Some(PanelAction::ExitSearch)
                }
                0x20 => Some(PanelAction::ToggleMark),
                0x43 if ctrl_pressed() => Some(PanelAction::ClipboardCopy),
                0x58 if ctrl_pressed() => Some(PanelAction::ClipboardCut),
                0x56 if ctrl_pressed() => Some(PanelAction::ClipboardPaste),
                0x4D if ctrl_pressed() => Some(PanelAction::ContextMenu),
                0x79 if shift_pressed() => Some(PanelAction::ContextMenu),
                0x71 => Some(PanelAction::Rename),
                0x74 => Some(PanelAction::Copy),
                0x75 => Some(PanelAction::Move),
                0x76 => Some(PanelAction::NewFolder),
                0x2E => Some(PanelAction::Delete),
                0x24 | 0x21 => Some(PanelAction::SelectFirst),
                0x23 | 0x22 => Some(PanelAction::SelectLast),
                0x7B => Some(PanelAction::Refresh),
                0x41 if ctrl_pressed() => Some(PanelAction::MarkAll),
                0x46 if ctrl_pressed() => Some(PanelAction::Search),
                0x25 | 0x27 => Some(PanelAction::SwitchPanel),
                _ => None,
            };

            if let Some(action) = handled {
                PostMessageW(
                    parent,
                    WM_PANEL_ACTION,
                    action.as_wparam(),
                    panel_index as LPARAM,
                );
                return 0;
            }
        }
        WM_COPY_MSG | WM_CUT_MSG | WM_PASTE_MSG => {
            let action = match msg {
                WM_COPY_MSG => PanelAction::ClipboardCopy,
                WM_CUT_MSG => PanelAction::ClipboardCut,
                WM_PASTE_MSG => PanelAction::ClipboardPaste,
                _ => unreachable!(),
            };
            PostMessageW(
                parent,
                WM_PANEL_ACTION,
                action.as_wparam(),
                panel_index as LPARAM,
            );
            return 0;
        }
        0x0102 => {
            let clipboard_action = match wparam as u32 {
                0x03 => Some(PanelAction::ClipboardCopy),
                0x18 => Some(PanelAction::ClipboardCut),
                0x16 => Some(PanelAction::ClipboardPaste),
                _ => None,
            };
            if let Some(action) = clipboard_action {
                PostMessageW(
                    parent,
                    WM_PANEL_ACTION,
                    action.as_wparam(),
                    panel_index as LPARAM,
                );
                return 0;
            }
            if wparam as u32 == 0x20 {
                return 0;
            }
            if !ctrl_pressed() {
                if let Some(ch) = char::from_u32(wparam as u32) {
                    if ch.is_ascii_alphanumeric() {
                        PostMessageW(parent, WM_PANEL_SEARCH, ch as WPARAM, panel_index as LPARAM);
                        return 0;
                    }
                }
            }
        }
        WM_NCDESTROY => {
            RemoveWindowSubclass(hwnd, Some(listbox_subclass_proc), subclass_id);
        }
        _ => {}
    }

    DefSubclassProc(hwnd, msg, wparam, lparam)
}

unsafe extern "system" fn input_dialog_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_NCCREATE => {
            let create =
                lparam as *const windows_sys::Win32::UI::WindowsAndMessaging::CREATESTRUCTW;
            let ptr = (*create).lpCreateParams as *mut InputDialogState;
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, ptr as isize);
            1
        }
        WM_CREATE => {
            let Some(state) = input_dialog_state_mut(hwnd) else {
                return -1;
            };

            let font = GetStockObject(17) as HFONT;
            let prompt_lines = vec![format!(
                "Uzupełnij pole {}.",
                state.prompt.trim().trim_end_matches(':')
            )];
            let prompt_height = dialog_list_height(prompt_lines.len());
            let prompt_hwnd = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                wide("LISTBOX").as_ptr(),
                wide("").as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_VSCROLL | LBS_NOTIFY as u32,
                12,
                14,
                360,
                prompt_height,
                hwnd,
                ID_DIALOG_INFO as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );
            let label_hwnd = CreateWindowExW(
                0,
                wide("STATIC").as_ptr(),
                wide(&make_access_label(&state.prompt)).as_ptr(),
                WS_CHILD | WS_VISIBLE,
                12,
                22 + prompt_height,
                140,
                20,
                hwnd,
                null_mut(),
                GetModuleHandleW(null()),
                null_mut(),
            );
            let edit_hwnd = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                wide("EDIT").as_ptr(),
                wide(&state.initial).as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_BORDER,
                158,
                18 + prompt_height,
                214,
                26,
                hwnd,
                ID_DIALOG_EDIT as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );
            let ok_hwnd = CreateWindowExW(
                0,
                wide("BUTTON").as_ptr(),
                wide("OK").as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON as u32,
                196,
                58 + prompt_height,
                84,
                28,
                hwnd,
                ID_DIALOG_OK as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );
            let cancel_hwnd = CreateWindowExW(
                0,
                wide("BUTTON").as_ptr(),
                wide("Anuluj").as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON as u32,
                288,
                58 + prompt_height,
                84,
                28,
                hwnd,
                ID_DIALOG_CANCEL as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );

            SendMessageW(prompt_hwnd, 0x0030, font as WPARAM, 1);
            SendMessageW(label_hwnd, 0x0030, font as WPARAM, 1);
            SendMessageW(edit_hwnd, 0x0030, font as WPARAM, 1);
            SendMessageW(ok_hwnd, 0x0030, font as WPARAM, 1);
            SendMessageW(cancel_hwnd, 0x0030, font as WPARAM, 1);
            SetWindowSubclass(edit_hwnd, Some(edit_subclass_proc), 1, 0);
            SetWindowSubclass(ok_hwnd, Some(dialog_button_subclass_proc), 4, 0);
            SetWindowSubclass(cancel_hwnd, Some(dialog_button_subclass_proc), 5, 0);
            SetWindowSubclass(
                prompt_hwnd,
                Some(dialog_listbox_subclass_proc),
                2,
                pack_dialog_button_ids(ID_DIALOG_OK, ID_DIALOG_CANCEL),
            );
            populate_listbox_lines(prompt_hwnd, &prompt_lines);

            state.prompt_hwnd = prompt_hwnd;
            state.edit_hwnd = edit_hwnd;
            state.ok_hwnd = ok_hwnd;
            state.cancel_hwnd = cancel_hwnd;
            state.prompt_lines = prompt_lines;
            SetFocus(edit_hwnd);
            0
        }
        WM_DIALOG_NAVIGATE => {
            let Some(state) = input_dialog_state_mut(hwnd) else {
                return 0;
            };
            let order = [
                state.prompt_hwnd,
                state.edit_hwnd,
                state.ok_hwnd,
                state.cancel_hwnd,
            ];
            focus_in_order(&order, lparam as HWND, wparam != 0);
            0
        }
        WM_COMMAND => {
            let id = loword(wparam as u32) as i32;
            let code = hiword(wparam as u32);
            if id == ID_DIALOG_INFO && code == LBN_SELCHANGE as u16 {
                if let Some(state) = input_dialog_state_mut(hwnd) {
                    speak_dialog_list_selection(state.owner, lparam as HWND, &state.prompt_lines);
                    return 0;
                }
            }
            match id {
                ID_DIALOG_OK => {
                    if let Some(state) = input_dialog_state_mut(hwnd) {
                        state.result = Some(read_window_text(state.edit_hwnd));
                        state.accepted = true;
                        DestroyWindow(hwnd);
                    }
                    0
                }
                ID_DIALOG_CANCEL => {
                    if let Some(state) = input_dialog_state_mut(hwnd) {
                        state.result = None;
                        state.accepted = false;
                        DestroyWindow(hwnd);
                    }
                    0
                }
                _ => DefWindowProcW(hwnd, msg, wparam, lparam),
            }
        }
        WM_CLOSE => {
            if let Some(state) = input_dialog_state_mut(hwnd) {
                state.result = None;
                state.accepted = false;
            }
            DestroyWindow(hwnd);
            0
        }
        WM_CTLCOLORSTATIC | WM_CTLCOLOREDIT | WM_CTLCOLORLISTBOX => {
            let hdc = wparam as HDC;
            SetTextColor(hdc, YELLOW);
            SetBkColor(hdc, BLACK);
            if msg == WM_CTLCOLORSTATIC {
                SetBkMode(hdc, TRANSPARENT as i32);
            }
            GetStockObject(4) as LRESULT
        }
        WM_DESTROY => {
            if let Some(state) = input_dialog_state_mut(hwnd) {
                state.done = true;
            }
            0
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

unsafe extern "system" fn search_dialog_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_NCCREATE => {
            let create =
                lparam as *const windows_sys::Win32::UI::WindowsAndMessaging::CREATESTRUCTW;
            let ptr = (*create).lpCreateParams as *mut SearchDialogState;
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, ptr as isize);
            1
        }
        WM_CREATE => {
            let Some(state) = search_dialog_state_mut(hwnd) else {
                return -1;
            };

            let font = GetStockObject(17) as HFONT;
            let prompt_height = dialog_list_height(state.prompt_lines.len());
            let prompt_hwnd = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                wide("LISTBOX").as_ptr(),
                wide("").as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_VSCROLL | LBS_NOTIFY as u32,
                12,
                14,
                420,
                prompt_height,
                hwnd,
                ID_DIALOG_INFO as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );
            let label_hwnd = CreateWindowExW(
                0,
                wide("STATIC").as_ptr(),
                wide("&Szukaj:").as_ptr(),
                WS_CHILD | WS_VISIBLE,
                12,
                22 + prompt_height,
                120,
                20,
                hwnd,
                null_mut(),
                GetModuleHandleW(null()),
                null_mut(),
            );
            let edit_hwnd = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                wide("EDIT").as_ptr(),
                wide("").as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_BORDER,
                136,
                18 + prompt_height,
                296,
                26,
                hwnd,
                ID_DIALOG_EDIT as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );
            let local_hwnd = CreateWindowExW(
                0,
                wide("BUTTON").as_ptr(),
                wide("Szukaj w katalogu").as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON as u32,
                72,
                58 + prompt_height,
                132,
                28,
                hwnd,
                ID_DIALOG_SEARCH_LOCAL as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );
            let recursive_hwnd = CreateWindowExW(
                0,
                wide("BUTTON").as_ptr(),
                wide("Szukaj rekurencyjnie").as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON as u32,
                212,
                58 + prompt_height,
                148,
                28,
                hwnd,
                ID_DIALOG_SEARCH_RECURSIVE as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );
            let cancel_hwnd = CreateWindowExW(
                0,
                wide("BUTTON").as_ptr(),
                wide("Anuluj").as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON as u32,
                368,
                58 + prompt_height,
                84,
                28,
                hwnd,
                ID_DIALOG_CANCEL as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );

            SendMessageW(prompt_hwnd, 0x0030, font as WPARAM, 1);
            SendMessageW(label_hwnd, 0x0030, font as WPARAM, 1);
            SendMessageW(edit_hwnd, 0x0030, font as WPARAM, 1);
            SendMessageW(local_hwnd, 0x0030, font as WPARAM, 1);
            SendMessageW(recursive_hwnd, 0x0030, font as WPARAM, 1);
            SendMessageW(cancel_hwnd, 0x0030, font as WPARAM, 1);
            SetWindowSubclass(edit_hwnd, Some(search_edit_subclass_proc), 6, 0);
            SetWindowSubclass(local_hwnd, Some(dialog_button_subclass_proc), 7, 0);
            SetWindowSubclass(recursive_hwnd, Some(dialog_button_subclass_proc), 8, 0);
            SetWindowSubclass(cancel_hwnd, Some(dialog_button_subclass_proc), 9, 0);
            SetWindowSubclass(
                prompt_hwnd,
                Some(dialog_listbox_subclass_proc),
                10,
                pack_dialog_button_ids(ID_DIALOG_SEARCH_LOCAL, ID_DIALOG_CANCEL),
            );
            populate_listbox_lines(prompt_hwnd, &state.prompt_lines);

            state.prompt_hwnd = prompt_hwnd;
            state.edit_hwnd = edit_hwnd;
            state.local_hwnd = local_hwnd;
            state.recursive_hwnd = recursive_hwnd;
            state.cancel_hwnd = cancel_hwnd;
            SetFocus(edit_hwnd);
            SendMessageW(edit_hwnd, EM_SETSEL, 0, -1isize as LPARAM);
            0
        }
        WM_DIALOG_NAVIGATE => {
            let Some(state) = search_dialog_state_mut(hwnd) else {
                return 0;
            };
            let order = [
                state.prompt_hwnd,
                state.edit_hwnd,
                state.local_hwnd,
                state.recursive_hwnd,
                state.cancel_hwnd,
            ];
            focus_in_order(&order, lparam as HWND, wparam != 0);
            0
        }
        WM_COMMAND => {
            let id = loword(wparam as u32) as i32;
            let code = hiword(wparam as u32);
            if id == ID_DIALOG_INFO && code == LBN_SELCHANGE as u16 {
                if let Some(state) = search_dialog_state_mut(hwnd) {
                    speak_dialog_list_selection(state.owner, lparam as HWND, &state.prompt_lines);
                    return 0;
                }
            }
            match id {
                ID_DIALOG_SEARCH_LOCAL | ID_DIALOG_SEARCH_RECURSIVE => {
                    if let Some(state) = search_dialog_state_mut(hwnd) {
                        let query = read_window_text(state.edit_hwnd);
                        state.result = Some((query, id == ID_DIALOG_SEARCH_RECURSIVE));
                        state.accepted = true;
                        DestroyWindow(hwnd);
                    }
                    0
                }
                ID_DIALOG_CANCEL => {
                    if let Some(state) = search_dialog_state_mut(hwnd) {
                        state.result = None;
                        state.accepted = false;
                        DestroyWindow(hwnd);
                    }
                    0
                }
                _ => DefWindowProcW(hwnd, msg, wparam, lparam),
            }
        }
        WM_CLOSE => {
            if let Some(state) = search_dialog_state_mut(hwnd) {
                state.result = None;
                state.accepted = false;
            }
            DestroyWindow(hwnd);
            0
        }
        WM_CTLCOLORSTATIC | WM_CTLCOLOREDIT | WM_CTLCOLORLISTBOX => {
            let hdc = wparam as HDC;
            SetTextColor(hdc, YELLOW);
            SetBkColor(hdc, BLACK);
            if msg == WM_CTLCOLORSTATIC {
                SetBkMode(hdc, TRANSPARENT as i32);
            }
            GetStockObject(4) as LRESULT
        }
        WM_DESTROY => {
            if let Some(state) = search_dialog_state_mut(hwnd) {
                state.done = true;
            }
            0
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

unsafe extern "system" fn network_dialog_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_NCCREATE => {
            let create =
                lparam as *const windows_sys::Win32::UI::WindowsAndMessaging::CREATESTRUCTW;
            let ptr = (*create).lpCreateParams as *mut NetworkDialogState;
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, ptr as isize);
            1
        }
        WM_CREATE => {
            let Some(state) = network_dialog_state_mut(hwnd) else {
                return -1;
            };
            let font = GetStockObject(17) as HFONT;
            let prompt_height = dialog_list_height(state.prompt_lines.len());
            let prompt_hwnd = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                wide("LISTBOX").as_ptr(),
                wide("").as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_VSCROLL | LBS_NOTIFY as u32,
                12,
                12,
                520,
                prompt_height,
                hwnd,
                ID_DIALOG_INFO as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );
            SendMessageW(prompt_hwnd, 0x0030, font as WPARAM, 1);
            populate_listbox_lines(prompt_hwnd, &state.prompt_lines);
            SetWindowSubclass(
                prompt_hwnd,
                Some(dialog_listbox_subclass_proc),
                21,
                pack_dialog_button_ids(ID_DIALOG_OK, ID_DIALOG_CANCEL),
            );

            let field_left = 12;
            let label_width = 170;
            let edit_left = 190;
            let edit_width = 342;
            let mut row = 22 + prompt_height;
            let row_height = 30;

            let protocol_label = CreateWindowExW(
                0,
                wide("STATIC").as_ptr(),
                wide(&make_access_label("Typ zasobu:")).as_ptr(),
                WS_CHILD | WS_VISIBLE,
                field_left,
                row + 4,
                label_width,
                20,
                hwnd,
                null_mut(),
                GetModuleHandleW(null()),
                null_mut(),
            );
            let protocol_hwnd = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                wide("COMBOBOX").as_ptr(),
                wide("").as_ptr(),
                WS_CHILD
                    | WS_VISIBLE
                    | WS_TABSTOP
                    | WS_VSCROLL
                    | WS_GROUP_STYLE
                    | CBS_DROPDOWNLIST as u32,
                edit_left,
                row,
                edit_width,
                150,
                hwnd,
                ID_NET_PROTOCOL as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );
            for protocol in NetworkProtocol::ALL {
                SendMessageW(
                    protocol_hwnd,
                    CB_ADDSTRING,
                    0,
                    wide(protocol.label()).as_ptr() as LPARAM,
                );
            }
            let selected_protocol = NetworkProtocol::ALL
                .iter()
                .position(|protocol| *protocol == state.initial.protocol)
                .unwrap_or(0);
            SendMessageW(protocol_hwnd, CB_SETCURSEL, selected_protocol, 0);
            row += row_height;

            let host_label = CreateWindowExW(
                0,
                wide("STATIC").as_ptr(),
                wide(&make_access_label("Adres hosta:")).as_ptr(),
                WS_CHILD | WS_VISIBLE,
                field_left,
                row + 4,
                label_width,
                20,
                hwnd,
                null_mut(),
                GetModuleHandleW(null()),
                null_mut(),
            );
            let host_hwnd = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                wide("EDIT").as_ptr(),
                wide(&state.initial.host).as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_BORDER,
                edit_left,
                row,
                edit_width,
                24,
                hwnd,
                ID_NET_HOST as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );
            row += row_height;

            let user_label = CreateWindowExW(
                0,
                wide("STATIC").as_ptr(),
                wide(&make_access_label("Nazwa użytkownika:")).as_ptr(),
                WS_CHILD | WS_VISIBLE,
                field_left,
                row + 4,
                label_width,
                20,
                hwnd,
                null_mut(),
                GetModuleHandleW(null()),
                null_mut(),
            );
            let username_hwnd = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                wide("EDIT").as_ptr(),
                wide(&state.initial.username).as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_BORDER,
                edit_left,
                row,
                edit_width,
                24,
                hwnd,
                ID_NET_USERNAME as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );
            row += row_height;

            let password_label = CreateWindowExW(
                0,
                wide("STATIC").as_ptr(),
                wide(&make_access_label("Hasło:")).as_ptr(),
                WS_CHILD | WS_VISIBLE,
                field_left,
                row + 4,
                label_width,
                20,
                hwnd,
                null_mut(),
                GetModuleHandleW(null()),
                null_mut(),
            );
            let password_hwnd = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                wide("EDIT").as_ptr(),
                wide(&state.initial.password).as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_BORDER | ES_PASSWORD,
                edit_left,
                row,
                edit_width,
                24,
                hwnd,
                ID_NET_PASSWORD as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );
            row += row_height;

            let anonymous_hwnd = CreateWindowExW(
                0,
                wide("BUTTON").as_ptr(),
                wide("Połącz bez danych logowania").as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX,
                edit_left,
                row,
                240,
                24,
                hwnd,
                ID_NET_ANONYMOUS as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );
            row += row_height;

            let key_label = CreateWindowExW(
                0,
                wide("STATIC").as_ptr(),
                wide(&make_access_label("Klucz SSH:")).as_ptr(),
                WS_CHILD | WS_VISIBLE,
                field_left,
                row + 4,
                label_width,
                20,
                hwnd,
                null_mut(),
                GetModuleHandleW(null()),
                null_mut(),
            );
            let ssh_key_hwnd = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                wide("EDIT").as_ptr(),
                wide(&state.initial.ssh_key).as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_BORDER,
                edit_left,
                row,
                246,
                24,
                hwnd,
                ID_NET_SSH_KEY as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );
            let ssh_key_browse_hwnd = CreateWindowExW(
                0,
                wide("BUTTON").as_ptr(),
                wide("Wybierz...").as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON as u32,
                edit_left + 256,
                row,
                86,
                24,
                hwnd,
                ID_NET_SSH_KEY_BROWSE as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );
            row += row_height;

            let directory_label = CreateWindowExW(
                0,
                wide("STATIC").as_ptr(),
                wide(&make_access_label("Domyślny katalog:")).as_ptr(),
                WS_CHILD | WS_VISIBLE,
                field_left,
                row + 4,
                label_width,
                20,
                hwnd,
                null_mut(),
                GetModuleHandleW(null()),
                null_mut(),
            );
            let directory_hwnd = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                wide("EDIT").as_ptr(),
                wide(&state.initial.default_directory).as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_BORDER,
                edit_left,
                row,
                edit_width,
                24,
                hwnd,
                ID_NET_DIRECTORY as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );
            row += row_height;

            let display_label = CreateWindowExW(
                0,
                wide("STATIC").as_ptr(),
                wide(&make_access_label("Nazwa wyświetlana:")).as_ptr(),
                WS_CHILD | WS_VISIBLE,
                field_left,
                row + 4,
                label_width,
                20,
                hwnd,
                null_mut(),
                GetModuleHandleW(null()),
                null_mut(),
            );
            let display_name_hwnd = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                wide("EDIT").as_ptr(),
                wide(&state.initial.display_name).as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_BORDER,
                edit_left,
                row,
                edit_width,
                24,
                hwnd,
                ID_NET_DISPLAY_NAME as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );
            row += row_height + 6;

            let ok_hwnd = CreateWindowExW(
                0,
                wide("BUTTON").as_ptr(),
                wide("Dodaj").as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON as u32,
                350,
                row,
                84,
                28,
                hwnd,
                ID_DIALOG_OK as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );
            let cancel_hwnd = CreateWindowExW(
                0,
                wide("BUTTON").as_ptr(),
                wide("Anuluj").as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON as u32,
                448,
                row,
                84,
                28,
                hwnd,
                ID_DIALOG_CANCEL as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );

            for control in [
                protocol_label,
                protocol_hwnd,
                host_label,
                host_hwnd,
                user_label,
                username_hwnd,
                password_label,
                password_hwnd,
                anonymous_hwnd,
                key_label,
                ssh_key_hwnd,
                ssh_key_browse_hwnd,
                directory_label,
                directory_hwnd,
                display_label,
                display_name_hwnd,
                ok_hwnd,
                cancel_hwnd,
            ] {
                SendMessageW(control, 0x0030, font as WPARAM, 1);
            }

            for (index, control) in [
                host_hwnd,
                username_hwnd,
                password_hwnd,
                ssh_key_hwnd,
                directory_hwnd,
                display_name_hwnd,
            ]
            .into_iter()
            .enumerate()
            {
                SetWindowSubclass(control, Some(edit_subclass_proc), 40 + index, 0);
            }
            SetWindowSubclass(anonymous_hwnd, Some(dialog_button_subclass_proc), 50, 0);
            SetWindowSubclass(
                ssh_key_browse_hwnd,
                Some(dialog_button_subclass_proc),
                51,
                0,
            );
            SetWindowSubclass(ok_hwnd, Some(dialog_button_subclass_proc), 52, 0);
            SetWindowSubclass(cancel_hwnd, Some(dialog_button_subclass_proc), 53, 0);

            state.info_hwnd = prompt_hwnd;
            state.protocol_hwnd = protocol_hwnd;
            state.host_hwnd = host_hwnd;
            state.username_hwnd = username_hwnd;
            state.password_hwnd = password_hwnd;
            state.ssh_key_hwnd = ssh_key_hwnd;
            state.ssh_key_browse_hwnd = ssh_key_browse_hwnd;
            state.directory_hwnd = directory_hwnd;
            state.display_name_hwnd = display_name_hwnd;
            state.anonymous_hwnd = anonymous_hwnd;
            state.ok_hwnd = ok_hwnd;
            state.cancel_hwnd = cancel_hwnd;

            set_radio_checked(anonymous_hwnd, state.initial.anonymous);
            update_network_dialog_state(hwnd);
            SetFocus(protocol_hwnd);
            0
        }
        WM_DIALOG_NAVIGATE => {
            let Some(state) = network_dialog_state_mut(hwnd) else {
                return 0;
            };
            let mut order = Vec::with_capacity(12);
            order.push(state.info_hwnd);
            order.push(state.protocol_hwnd);
            order.push(state.host_hwnd);
            order.push(state.username_hwnd);
            order.push(state.password_hwnd);
            order.push(state.anonymous_hwnd);
            order.push(state.ssh_key_hwnd);
            order.push(state.ssh_key_browse_hwnd);
            order.push(state.directory_hwnd);
            order.push(state.display_name_hwnd);
            order.push(state.ok_hwnd);
            order.push(state.cancel_hwnd);
            focus_in_order(&order, lparam as HWND, wparam != 0);
            0
        }
        WM_COMMAND => {
            let id = loword(wparam as u32) as i32;
            let code = hiword(wparam as u32);
            if id == ID_DIALOG_INFO && code == LBN_SELCHANGE as u16 {
                if let Some(state) = network_dialog_state_mut(hwnd) {
                    speak_dialog_list_selection(state.owner, lparam as HWND, &state.prompt_lines);
                    return 0;
                }
            }
            if id == ID_NET_PROTOCOL && code == CBN_SELCHANGE as u16 {
                update_network_dialog_state(hwnd);
                return 0;
            }
            match id {
                ID_NET_SSH_KEY_BROWSE => {
                    if let Some(state) = network_dialog_state_mut(hwnd) {
                        let current = read_window_text(state.ssh_key_hwnd);
                        if let Some(path) =
                            choose_file_dialog(hwnd, "Wybierz plik klucza SSH", &current)
                        {
                            let value = path.to_string_lossy().to_string();
                            SetWindowTextW(state.ssh_key_hwnd, wide(&value).as_ptr());
                            if let Some(app) = app_state_mut(state.owner) {
                                let message = format!("wybrano plik {}", display_name(&path));
                                app.nvda.speak(&message);
                            }
                        }
                    }
                    0
                }
                ID_DIALOG_OK => {
                    if let Some(state) = network_dialog_state_mut(hwnd) {
                        let host = read_window_text(state.host_hwnd);
                        if host.trim().is_empty() {
                            if let Some(app) = app_state_mut(state.owner) {
                                app.nvda.speak("adres hosta nie może być pusty");
                            }
                            SetFocus(state.host_hwnd);
                            return 0;
                        }
                        let resource = NetworkResource {
                            protocol: selected_network_protocol(state),
                            host: host.trim().to_string(),
                            username: read_window_text(state.username_hwnd).trim().to_string(),
                            password: read_window_text(state.password_hwnd),
                            root_password: state.initial.root_password.clone(),
                            sudo_password: state.initial.sudo_password.clone(),
                            ssh_key: read_window_text(state.ssh_key_hwnd).trim().to_string(),
                            default_directory: read_window_text(state.directory_hwnd)
                                .trim()
                                .to_string(),
                            display_name: read_window_text(state.display_name_hwnd)
                                .trim()
                                .to_string(),
                            anonymous: is_button_checked(state.anonymous_hwnd),
                        };
                        state.result = Some(resource);
                        state.accepted = true;
                    }
                    DestroyWindow(hwnd);
                    0
                }
                ID_DIALOG_CANCEL => {
                    DestroyWindow(hwnd);
                    0
                }
                _ => DefWindowProcW(hwnd, msg, wparam, lparam),
            }
        }
        WM_CTLCOLORSTATIC | WM_CTLCOLOREDIT | WM_CTLCOLORLISTBOX => {
            let hdc = wparam as HDC;
            SetTextColor(hdc, YELLOW);
            SetBkColor(hdc, BLACK);
            if msg == WM_CTLCOLORSTATIC {
                SetBkMode(hdc, TRANSPARENT as i32);
            }
            GetStockObject(4) as LRESULT
        }
        WM_CLOSE => {
            DestroyWindow(hwnd);
            0
        }
        WM_DESTROY => {
            if let Some(state) = network_dialog_state_mut(hwnd) {
                state.done = true;
            }
            0
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

unsafe extern "system" fn archive_create_dialog_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_NCCREATE => {
            let create =
                lparam as *const windows_sys::Win32::UI::WindowsAndMessaging::CREATESTRUCTW;
            let ptr = (*create).lpCreateParams as *mut ArchiveCreateDialogState;
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, ptr as isize);
            1
        }
        WM_CREATE => {
            let Some(state) = archive_create_dialog_state_mut(hwnd) else {
                return -1;
            };
            let font = GetStockObject(17) as HFONT;
            let prompt_height = dialog_list_height(state.prompt_lines.len());
            let info_hwnd = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                wide("LISTBOX").as_ptr(),
                wide("").as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_VSCROLL | LBS_NOTIFY as u32,
                12,
                12,
                520,
                prompt_height,
                hwnd,
                ID_DIALOG_INFO as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );
            populate_listbox_lines(info_hwnd, &state.prompt_lines);
            SetWindowSubclass(
                info_hwnd,
                Some(dialog_listbox_subclass_proc),
                71,
                pack_dialog_button_ids(ID_DIALOG_OK, ID_DIALOG_CANCEL),
            );

            let field_left = 12;
            let label_width = 178;
            let edit_left = 200;
            let edit_width = 332;
            let mut row = 22 + prompt_height;
            let row_height = 30;

            let format_label = CreateWindowExW(
                0,
                wide("STATIC").as_ptr(),
                wide(&make_access_label("Rodzaj archiwum lub obrazu:")).as_ptr(),
                WS_CHILD | WS_VISIBLE,
                field_left,
                row + 4,
                label_width,
                20,
                hwnd,
                null_mut(),
                GetModuleHandleW(null()),
                null_mut(),
            );
            let format_hwnd = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                wide("COMBOBOX").as_ptr(),
                wide("").as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_VSCROLL | CBS_DROPDOWNLIST as u32,
                edit_left,
                row,
                edit_width,
                140,
                hwnd,
                ID_ARCHIVE_TYPE as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );
            for format in ArchiveFormat::ALL {
                SendMessageW(
                    format_hwnd,
                    CB_ADDSTRING,
                    0,
                    wide(format.label()).as_ptr() as LPARAM,
                );
            }
            SendMessageW(format_hwnd, CB_SETCURSEL, 0, 0);
            row += row_height;

            let name_label = CreateWindowExW(
                0,
                wide("STATIC").as_ptr(),
                wide(&make_access_label("Nazwa pliku:")).as_ptr(),
                WS_CHILD | WS_VISIBLE,
                field_left,
                row + 4,
                label_width,
                20,
                hwnd,
                null_mut(),
                GetModuleHandleW(null()),
                null_mut(),
            );
            let name_hwnd = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                wide("EDIT").as_ptr(),
                wide(&state.initial_name).as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_BORDER,
                edit_left,
                row,
                edit_width,
                24,
                hwnd,
                ID_ARCHIVE_NAME as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );
            row += row_height;

            let level_label = CreateWindowExW(
                0,
                wide("STATIC").as_ptr(),
                wide(&make_access_label("Stopień kompresji:")).as_ptr(),
                WS_CHILD | WS_VISIBLE,
                field_left,
                row + 4,
                label_width,
                20,
                hwnd,
                null_mut(),
                GetModuleHandleW(null()),
                null_mut(),
            );
            let level_hwnd = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                wide("COMBOBOX").as_ptr(),
                wide("").as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_VSCROLL | CBS_DROPDOWNLIST as u32,
                edit_left,
                row,
                edit_width,
                140,
                hwnd,
                ID_ARCHIVE_LEVEL as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );
            for label in ["brak", "szybka", "normalna", "maksymalna"] {
                SendMessageW(level_hwnd, CB_ADDSTRING, 0, wide(label).as_ptr() as LPARAM);
            }
            SendMessageW(level_hwnd, CB_SETCURSEL, 2, 0);
            row += row_height;

            let encrypted_hwnd = CreateWindowExW(
                0,
                wide("BUTTON").as_ptr(),
                wide("Szyfruj archiwum").as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX,
                edit_left,
                row,
                220,
                24,
                hwnd,
                ID_ARCHIVE_ENCRYPTED as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );
            row += row_height;

            let encryption_label = CreateWindowExW(
                0,
                wide("STATIC").as_ptr(),
                wide(&make_access_label("Rodzaj szyfrowania:")).as_ptr(),
                WS_CHILD | WS_VISIBLE,
                field_left,
                row + 4,
                label_width,
                20,
                hwnd,
                null_mut(),
                GetModuleHandleW(null()),
                null_mut(),
            );
            let encryption_hwnd = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                wide("COMBOBOX").as_ptr(),
                wide("").as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_VSCROLL | CBS_DROPDOWNLIST as u32,
                edit_left,
                row,
                edit_width,
                120,
                hwnd,
                ID_ARCHIVE_ENCRYPTION as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );
            for label in ["AES-256", "ZipCrypto"] {
                SendMessageW(
                    encryption_hwnd,
                    CB_ADDSTRING,
                    0,
                    wide(label).as_ptr() as LPARAM,
                );
            }
            SendMessageW(encryption_hwnd, CB_SETCURSEL, 0, 0);
            row += row_height;

            let password_label = CreateWindowExW(
                0,
                wide("STATIC").as_ptr(),
                wide(&make_access_label("Hasło:")).as_ptr(),
                WS_CHILD | WS_VISIBLE,
                field_left,
                row + 4,
                label_width,
                20,
                hwnd,
                null_mut(),
                GetModuleHandleW(null()),
                null_mut(),
            );
            let password_hwnd = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                wide("EDIT").as_ptr(),
                wide("").as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_BORDER | ES_PASSWORD,
                edit_left,
                row,
                edit_width,
                24,
                hwnd,
                ID_ARCHIVE_PASSWORD as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );
            row += row_height;

            let volume_label = CreateWindowExW(
                0,
                wide("STATIC").as_ptr(),
                wide(&make_access_label("Rozmiar części:")).as_ptr(),
                WS_CHILD | WS_VISIBLE,
                field_left,
                row + 4,
                label_width,
                20,
                hwnd,
                null_mut(),
                GetModuleHandleW(null()),
                null_mut(),
            );
            let volume_hwnd = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                wide("EDIT").as_ptr(),
                wide("").as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_BORDER,
                edit_left,
                row,
                edit_width,
                24,
                hwnd,
                ID_ARCHIVE_VOLUME as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );
            row += row_height + 6;

            let ok_hwnd = CreateWindowExW(
                0,
                wide("BUTTON").as_ptr(),
                wide("Utwórz").as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON as u32,
                350,
                row,
                84,
                28,
                hwnd,
                ID_DIALOG_OK as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );
            let cancel_hwnd = CreateWindowExW(
                0,
                wide("BUTTON").as_ptr(),
                wide("Anuluj").as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON as u32,
                448,
                row,
                84,
                28,
                hwnd,
                ID_DIALOG_CANCEL as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );

            for control in [
                info_hwnd,
                format_label,
                format_hwnd,
                name_label,
                name_hwnd,
                level_label,
                level_hwnd,
                encrypted_hwnd,
                encryption_label,
                encryption_hwnd,
                password_label,
                password_hwnd,
                volume_label,
                volume_hwnd,
                ok_hwnd,
                cancel_hwnd,
            ] {
                SendMessageW(control, 0x0030, font as WPARAM, 1);
            }
            for (index, control) in [name_hwnd, password_hwnd, volume_hwnd]
                .into_iter()
                .enumerate()
            {
                SetWindowSubclass(control, Some(edit_subclass_proc), 80 + index, 0);
            }
            for (index, control) in [encrypted_hwnd, ok_hwnd, cancel_hwnd]
                .into_iter()
                .enumerate()
            {
                SetWindowSubclass(control, Some(dialog_button_subclass_proc), 90 + index, 0);
            }

            state.info_hwnd = info_hwnd;
            state.format_hwnd = format_hwnd;
            state.name_hwnd = name_hwnd;
            state.level_hwnd = level_hwnd;
            state.encrypted_hwnd = encrypted_hwnd;
            state.encryption_hwnd = encryption_hwnd;
            state.password_hwnd = password_hwnd;
            state.volume_hwnd = volume_hwnd;
            state.ok_hwnd = ok_hwnd;
            state.cancel_hwnd = cancel_hwnd;
            SetFocus(format_hwnd);
            0
        }
        WM_DIALOG_NAVIGATE => {
            let Some(state) = archive_create_dialog_state_mut(hwnd) else {
                return 0;
            };
            let order = [
                state.info_hwnd,
                state.format_hwnd,
                state.name_hwnd,
                state.level_hwnd,
                state.encrypted_hwnd,
                state.encryption_hwnd,
                state.password_hwnd,
                state.volume_hwnd,
                state.ok_hwnd,
                state.cancel_hwnd,
            ];
            focus_in_order(&order, lparam as HWND, wparam != 0);
            0
        }
        WM_COMMAND => {
            let id = loword(wparam as u32) as i32;
            match id {
                ID_DIALOG_OK => {
                    if let Some(state) = archive_create_dialog_state_mut(hwnd) {
                        let name = read_window_text(state.name_hwnd);
                        if name.trim().is_empty() {
                            if let Some(app) = app_state_mut(state.owner) {
                                app.nvda.speak("nazwa pliku nie może być pusta");
                            }
                            SetFocus(state.name_hwnd);
                            return 0;
                        }
                        let format_index =
                            SendMessageW(state.format_hwnd, CB_GETCURSEL, 0, 0).max(0) as usize;
                        let level_index =
                            SendMessageW(state.level_hwnd, CB_GETCURSEL, 0, 0).max(0) as usize;
                        let level = [0u8, 1, 5, 9].get(level_index).copied().unwrap_or(5);
                        let encryption_index =
                            SendMessageW(state.encryption_hwnd, CB_GETCURSEL, 0, 0).max(0);
                        let encryption = if encryption_index == 1 {
                            "ZipCrypto"
                        } else {
                            "AES-256"
                        };
                        state.result = Some(ArchiveCreateOptions {
                            format: ArchiveFormat::from_index(format_index),
                            name: name.trim().to_string(),
                            compression_level: level,
                            encrypted: is_button_checked(state.encrypted_hwnd),
                            encryption: encryption.to_string(),
                            password: read_window_text(state.password_hwnd),
                            volume_size: read_window_text(state.volume_hwnd).trim().to_string(),
                        });
                        state.accepted = true;
                    }
                    DestroyWindow(hwnd);
                    0
                }
                ID_DIALOG_CANCEL => {
                    DestroyWindow(hwnd);
                    0
                }
                _ => DefWindowProcW(hwnd, msg, wparam, lparam),
            }
        }
        WM_CTLCOLORSTATIC | WM_CTLCOLOREDIT | WM_CTLCOLORLISTBOX => {
            let hdc = wparam as HDC;
            SetTextColor(hdc, YELLOW);
            SetBkColor(hdc, BLACK);
            if msg == WM_CTLCOLORSTATIC {
                SetBkMode(hdc, TRANSPARENT as i32);
            }
            GetStockObject(4) as LRESULT
        }
        WM_CLOSE => {
            DestroyWindow(hwnd);
            0
        }
        WM_DESTROY => {
            if let Some(state) = archive_create_dialog_state_mut(hwnd) {
                state.done = true;
            }
            0
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

unsafe extern "system" fn discovery_dialog_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_NCCREATE => {
            let create =
                lparam as *const windows_sys::Win32::UI::WindowsAndMessaging::CREATESTRUCTW;
            let ptr = (*create).lpCreateParams as *mut DiscoveryDialogState;
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, ptr as isize);
            1
        }
        WM_CREATE => {
            let Some(state) = discovery_dialog_state_mut(hwnd) else {
                return -1;
            };
            let font = GetStockObject(17) as HFONT;
            let prompt_height = dialog_list_height(state.prompt_lines.len());
            let info_hwnd = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                wide("LISTBOX").as_ptr(),
                wide("").as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_VSCROLL | LBS_NOTIFY as u32,
                12,
                12,
                460,
                prompt_height,
                hwnd,
                ID_DIALOG_INFO as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );
            let list_hwnd = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                wide("LISTBOX").as_ptr(),
                wide("").as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_VSCROLL | LBS_NOTIFY as u32,
                12,
                24 + prompt_height,
                460,
                180,
                hwnd,
                ID_NET_DISCOVERY_LIST as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );
            let add_hwnd = CreateWindowExW(
                0,
                wide("BUTTON").as_ptr(),
                wide("Dodaj").as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON as u32,
                290,
                218 + prompt_height,
                84,
                28,
                hwnd,
                ID_NET_DISCOVERY_ADD as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );
            let cancel_hwnd = CreateWindowExW(
                0,
                wide("BUTTON").as_ptr(),
                wide("Anuluj").as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON as u32,
                388,
                218 + prompt_height,
                84,
                28,
                hwnd,
                ID_DIALOG_CANCEL as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );

            for control in [info_hwnd, list_hwnd, add_hwnd, cancel_hwnd] {
                SendMessageW(control, 0x0030, font as WPARAM, 1);
            }
            populate_listbox_lines(info_hwnd, &state.prompt_lines);
            populate_listbox_lines(
                list_hwnd,
                &state
                    .servers
                    .iter()
                    .map(DiscoveredServer::display_line)
                    .collect::<Vec<_>>(),
            );
            SendMessageW(list_hwnd, LB_SETCURSEL, 0, 0);
            SetWindowSubclass(
                info_hwnd,
                Some(dialog_listbox_subclass_proc),
                61,
                pack_dialog_button_ids(ID_NET_DISCOVERY_ADD, ID_DIALOG_CANCEL),
            );
            SetWindowSubclass(
                list_hwnd,
                Some(dialog_listbox_subclass_proc),
                62,
                pack_dialog_button_ids(ID_NET_DISCOVERY_ADD, ID_DIALOG_CANCEL),
            );
            SetWindowSubclass(add_hwnd, Some(dialog_button_subclass_proc), 63, 0);
            SetWindowSubclass(cancel_hwnd, Some(dialog_button_subclass_proc), 64, 0);
            state.info_hwnd = info_hwnd;
            state.list_hwnd = list_hwnd;
            state.add_hwnd = add_hwnd;
            state.cancel_hwnd = cancel_hwnd;
            SetFocus(list_hwnd);
            if let Some(server) = state.servers.first() {
                if let Some(app) = app_state_mut(state.owner) {
                    app.nvda.speak(&server.display_line());
                }
            }
            0
        }
        WM_DIALOG_NAVIGATE => {
            let Some(state) = discovery_dialog_state_mut(hwnd) else {
                return 0;
            };
            let order = [
                state.info_hwnd,
                state.list_hwnd,
                state.add_hwnd,
                state.cancel_hwnd,
            ];
            focus_in_order(&order, lparam as HWND, wparam != 0);
            0
        }
        WM_COMMAND => {
            let id = loword(wparam as u32) as i32;
            let code = hiword(wparam as u32);
            match id {
                ID_DIALOG_INFO if code == LBN_SELCHANGE as u16 => {
                    if let Some(state) = discovery_dialog_state_mut(hwnd) {
                        speak_dialog_list_selection(
                            state.owner,
                            lparam as HWND,
                            &state.prompt_lines,
                        );
                    }
                    0
                }
                ID_NET_DISCOVERY_LIST if code == LBN_SELCHANGE as u16 => {
                    if let Some(state) = discovery_dialog_state_mut(hwnd) {
                        let index = SendMessageW(state.list_hwnd, LB_GETCURSEL, 0, 0);
                        if index >= 0 {
                            if let Some(server) = state.servers.get(index as usize) {
                                if let Some(app) = app_state_mut(state.owner) {
                                    app.nvda.speak(&server.display_line());
                                }
                            }
                        }
                    }
                    0
                }
                ID_NET_DISCOVERY_ADD => {
                    if let Some(state) = discovery_dialog_state_mut(hwnd) {
                        let index = SendMessageW(state.list_hwnd, LB_GETCURSEL, 0, 0);
                        if index >= 0 {
                            state.result = state.servers.get(index as usize).cloned();
                            DestroyWindow(hwnd);
                            return 0;
                        }
                    }
                    0
                }
                ID_DIALOG_CANCEL => {
                    DestroyWindow(hwnd);
                    0
                }
                _ => DefWindowProcW(hwnd, msg, wparam, lparam),
            }
        }
        WM_CTLCOLORSTATIC | WM_CTLCOLOREDIT | WM_CTLCOLORLISTBOX => {
            let hdc = wparam as HDC;
            SetTextColor(hdc, YELLOW);
            SetBkColor(hdc, BLACK);
            if msg == WM_CTLCOLORSTATIC {
                SetBkMode(hdc, TRANSPARENT as i32);
            }
            GetStockObject(4) as LRESULT
        }
        WM_CLOSE => {
            DestroyWindow(hwnd);
            0
        }
        WM_DESTROY => {
            if let Some(state) = discovery_dialog_state_mut(hwnd) {
                state.done = true;
            }
            0
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

unsafe extern "system" fn edit_subclass_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
    subclass_id: usize,
    _ref_data: usize,
) -> LRESULT {
    if msg == WM_SETFOCUS {
        SendMessageW(hwnd, EM_SETSEL, 0, -1isize as LPARAM);
    }
    if msg == WM_KEYDOWN && wparam as u32 == 0x41 && ctrl_pressed() {
        SendMessageW(hwnd, EM_SETSEL, 0, -1isize as LPARAM);
        return 0;
    }
    if msg == WM_KEYDOWN && wparam as u32 == 0x09 {
        let parent = windows_sys::Win32::UI::WindowsAndMessaging::GetParent(hwnd);
        SendMessageW(
            parent,
            WM_DIALOG_NAVIGATE,
            if shift_pressed() { 1 } else { 0 },
            hwnd as LPARAM,
        );
        return 0;
    }
    if msg == WM_KEYDOWN && wparam as u32 == 0x0D {
        let parent = windows_sys::Win32::UI::WindowsAndMessaging::GetParent(hwnd);
        SendMessageW(parent, WM_COMMAND, ID_DIALOG_OK as WPARAM, 0);
        return 0;
    }
    if msg == WM_KEYDOWN && wparam as u32 == 0x1B {
        let parent = windows_sys::Win32::UI::WindowsAndMessaging::GetParent(hwnd);
        SendMessageW(parent, WM_COMMAND, ID_DIALOG_CANCEL as WPARAM, 0);
        return 0;
    }
    if msg == WM_NCDESTROY {
        RemoveWindowSubclass(hwnd, Some(edit_subclass_proc), subclass_id);
    }
    DefSubclassProc(hwnd, msg, wparam, lparam)
}

unsafe extern "system" fn search_edit_subclass_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
    subclass_id: usize,
    _ref_data: usize,
) -> LRESULT {
    if msg == WM_SETFOCUS {
        SendMessageW(hwnd, EM_SETSEL, 0, -1isize as LPARAM);
    }
    if msg == WM_KEYDOWN && wparam as u32 == 0x41 && ctrl_pressed() {
        SendMessageW(hwnd, EM_SETSEL, 0, -1isize as LPARAM);
        return 0;
    }
    if msg == WM_KEYDOWN && wparam as u32 == 0x09 {
        let parent = windows_sys::Win32::UI::WindowsAndMessaging::GetParent(hwnd);
        SendMessageW(
            parent,
            WM_DIALOG_NAVIGATE,
            if shift_pressed() { 1 } else { 0 },
            hwnd as LPARAM,
        );
        return 0;
    }
    if msg == WM_KEYDOWN && wparam as u32 == 0x0D {
        let parent = windows_sys::Win32::UI::WindowsAndMessaging::GetParent(hwnd);
        SendMessageW(parent, WM_COMMAND, ID_DIALOG_SEARCH_LOCAL as WPARAM, 0);
        return 0;
    }
    if msg == WM_KEYDOWN && wparam as u32 == 0x1B {
        let parent = windows_sys::Win32::UI::WindowsAndMessaging::GetParent(hwnd);
        SendMessageW(parent, WM_COMMAND, ID_DIALOG_CANCEL as WPARAM, 0);
        return 0;
    }
    if msg == WM_NCDESTROY {
        RemoveWindowSubclass(hwnd, Some(search_edit_subclass_proc), subclass_id);
    }
    DefSubclassProc(hwnd, msg, wparam, lparam)
}

unsafe extern "system" fn dialog_listbox_subclass_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
    subclass_id: usize,
    ref_data: usize,
) -> LRESULT {
    let default_id = (ref_data >> 16) as i32;
    let cancel_id = (ref_data & 0xffff) as i32;
    match msg {
        WM_GETDLGCODE => {
            return DefSubclassProc(hwnd, msg, wparam, lparam) | 0x0002;
        }
        WM_KEYDOWN if wparam as u32 == 0x09 => {
            let parent = windows_sys::Win32::UI::WindowsAndMessaging::GetParent(hwnd);
            SendMessageW(
                parent,
                WM_DIALOG_NAVIGATE,
                if shift_pressed() { 1 } else { 0 },
                hwnd as LPARAM,
            );
            return 0;
        }
        WM_KEYDOWN if wparam as u32 == 0x0D => {
            let parent = windows_sys::Win32::UI::WindowsAndMessaging::GetParent(hwnd);
            SendMessageW(parent, WM_COMMAND, default_id as WPARAM, 0);
            return 0;
        }
        WM_KEYDOWN if wparam as u32 == 0x1B && cancel_id != 0 => {
            let parent = windows_sys::Win32::UI::WindowsAndMessaging::GetParent(hwnd);
            SendMessageW(parent, WM_COMMAND, cancel_id as WPARAM, 0);
            return 0;
        }
        WM_NCDESTROY => {
            RemoveWindowSubclass(hwnd, Some(dialog_listbox_subclass_proc), subclass_id);
        }
        _ => {}
    }
    DefSubclassProc(hwnd, msg, wparam, lparam)
}

unsafe extern "system" fn dialog_button_subclass_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
    subclass_id: usize,
    _ref_data: usize,
) -> LRESULT {
    match msg {
        WM_GETDLGCODE => {
            return DefSubclassProc(hwnd, msg, wparam, lparam) | 0x0002;
        }
        WM_SETFOCUS => {
            let parent = windows_sys::Win32::UI::WindowsAndMessaging::GetParent(hwnd);
            if let Some(state) = operation_prompt_state_mut(parent) {
                speak_control_text(state.owner, hwnd);
            } else if let Some(state) = input_dialog_state_mut(parent) {
                speak_control_text(state.owner, hwnd);
            } else if let Some(state) = search_dialog_state_mut(parent) {
                speak_control_text(state.owner, hwnd);
            } else if let Some(state) = network_dialog_state_mut(parent) {
                speak_control_text(state.owner, hwnd);
            } else if let Some(state) = discovery_dialog_state_mut(parent) {
                speak_control_text(state.owner, hwnd);
            }
        }
        WM_KEYDOWN if wparam as u32 == 0x09 => {
            let parent = windows_sys::Win32::UI::WindowsAndMessaging::GetParent(hwnd);
            SendMessageW(
                parent,
                WM_DIALOG_NAVIGATE,
                if shift_pressed() { 1 } else { 0 },
                hwnd as LPARAM,
            );
            return 0;
        }
        WM_KEYDOWN if wparam as u32 == 0x0D => {
            let parent = windows_sys::Win32::UI::WindowsAndMessaging::GetParent(hwnd);
            let id = GetDlgCtrlID(hwnd);
            SendMessageW(parent, WM_COMMAND, id as WPARAM, hwnd as LPARAM);
            return 0;
        }
        WM_KEYDOWN if wparam as u32 == 0x1B => {
            let parent = windows_sys::Win32::UI::WindowsAndMessaging::GetParent(hwnd);
            SendMessageW(
                parent,
                WM_COMMAND,
                ID_DIALOG_CANCEL as WPARAM,
                hwnd as LPARAM,
            );
            return 0;
        }
        WM_NCDESTROY => {
            RemoveWindowSubclass(hwnd, Some(dialog_button_subclass_proc), subclass_id);
        }
        _ => {}
    }
    DefSubclassProc(hwnd, msg, wparam, lparam)
}

unsafe extern "system" fn progress_button_subclass_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
    subclass_id: usize,
    _ref_data: usize,
) -> LRESULT {
    match msg {
        WM_GETDLGCODE => {
            return DefSubclassProc(hwnd, msg, wparam, lparam) | 0x0003;
        }
        WM_SETFOCUS => {
            let parent = windows_sys::Win32::UI::WindowsAndMessaging::GetParent(hwnd);
            if let Some(state) = progress_dialog_state_mut(parent) {
                speak_control_text(state.owner, hwnd);
            }
        }
        WM_KEYDOWN if wparam as u32 == 0x09 => {
            let parent = windows_sys::Win32::UI::WindowsAndMessaging::GetParent(hwnd);
            SendMessageW(
                parent,
                WM_DIALOG_NAVIGATE,
                if shift_pressed() { 1 } else { 0 },
                hwnd as LPARAM,
            );
            return 0;
        }
        WM_KEYDOWN if matches!(wparam as u32, 0x26 | 0x28) => {
            let parent = windows_sys::Win32::UI::WindowsAndMessaging::GetParent(hwnd);
            let direction = if wparam as u32 == 0x26 { 0 } else { 1 };
            SendMessageW(parent, WM_PROGRESS_NAVIGATE, direction as WPARAM, 0);
            return 0;
        }
        WM_KEYDOWN if wparam as u32 == 0x0D => {
            let parent = windows_sys::Win32::UI::WindowsAndMessaging::GetParent(hwnd);
            SendMessageW(
                parent,
                WM_COMMAND,
                ID_DIALOG_CANCEL as WPARAM,
                hwnd as LPARAM,
            );
            return 0;
        }
        WM_KEYDOWN if wparam as u32 == 0x1B => {
            let parent = windows_sys::Win32::UI::WindowsAndMessaging::GetParent(hwnd);
            SendMessageW(
                parent,
                WM_COMMAND,
                ID_DIALOG_CANCEL as WPARAM,
                hwnd as LPARAM,
            );
            return 0;
        }
        WM_NCDESTROY => {
            RemoveWindowSubclass(hwnd, Some(progress_button_subclass_proc), subclass_id);
        }
        _ => {}
    }
    DefSubclassProc(hwnd, msg, wparam, lparam)
}

unsafe fn speak_progress_line(owner: HWND, lines: &[String], index: usize) {
    if let Some(line) = lines.get(index) {
        if let Some(app) = app_state_mut(owner) {
            app.nvda.speak(line);
        }
    }
}

unsafe fn speak_dialog_list_selection(owner: HWND, list_hwnd: HWND, lines: &[String]) {
    let index = SendMessageW(list_hwnd, LB_GETCURSEL, 0, 0);
    if index >= 0 {
        if let Some(line) = lines.get(index as usize) {
            if let Some(app) = app_state_mut(owner) {
                app.nvda.speak(line);
            }
        }
    }
}

unsafe fn speak_control_text(owner: HWND, control_hwnd: HWND) {
    let text = read_window_text(control_hwnd);
    if text.is_empty() {
        return;
    }
    if let Some(app) = app_state_mut(owner) {
        app.nvda.speak(&text);
    }
}

unsafe fn focus_in_order(order: &[HWND], current: HWND, previous: bool) -> bool {
    if order.is_empty() {
        return false;
    }
    let current_index = order.iter().position(|hwnd| *hwnd == current).unwrap_or(0);
    let next_index = if previous {
        current_index.checked_sub(1).unwrap_or(order.len() - 1)
    } else {
        (current_index + 1) % order.len()
    };
    let next = order[next_index];
    if next.is_null() {
        false
    } else {
        SetFocus(next);
        true
    }
}

unsafe fn set_radio_checked(control: HWND, checked: bool) {
    SendMessageW(
        control,
        BM_SETCHECK,
        if checked { BST_CHECKED } else { 0 },
        0,
    );
}

unsafe fn is_button_checked(control: HWND) -> bool {
    SendMessageW(control, BM_GETCHECK, 0, 0) as usize == BST_CHECKED
}

unsafe fn selected_network_protocol(state: &NetworkDialogState) -> NetworkProtocol {
    let selected = SendMessageW(state.protocol_hwnd, CB_GETCURSEL, 0, 0);
    if selected >= 0 {
        NetworkProtocol::ALL
            .get(selected as usize)
            .copied()
            .unwrap_or_default()
    } else {
        NetworkProtocol::default()
    }
}

unsafe fn update_network_dialog_state(hwnd: HWND) {
    let _ = hwnd;
}

unsafe extern "system" fn progress_dialog_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_NCCREATE => {
            let create =
                lparam as *const windows_sys::Win32::UI::WindowsAndMessaging::CREATESTRUCTW;
            let ptr = (*create).lpCreateParams as *mut ProgressDialogState;
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, ptr as isize);
            1
        }
        WM_CREATE => {
            let Some(state) = progress_dialog_state_mut(hwnd) else {
                return -1;
            };
            let font = GetStockObject(17) as HFONT;
            let text_height = dialog_list_height(state.lines.len());
            let info_hwnd = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                wide("STATIC").as_ptr(),
                wide("").as_ptr(),
                WS_CHILD | WS_VISIBLE,
                12,
                12,
                500,
                text_height,
                hwnd,
                ID_DIALOG_INFO as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );
            let action_hwnd = CreateWindowExW(
                0,
                wide("BUTTON").as_ptr(),
                wide("Anuluj").as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON as u32,
                416,
                24 + text_height,
                96,
                28,
                hwnd,
                ID_DIALOG_CANCEL as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );
            SendMessageW(info_hwnd, 0x0030, font as WPARAM, 1);
            SendMessageW(action_hwnd, 0x0030, font as WPARAM, 1);
            set_progress_text(info_hwnd, &state.lines, state.current_line);
            SetWindowSubclass(action_hwnd, Some(progress_button_subclass_proc), 1, 0);
            state.info_hwnd = info_hwnd;
            state.action_hwnd = action_hwnd;
            SetFocus(action_hwnd);
            speak_control_text(state.owner, action_hwnd);
            0
        }
        WM_DIALOG_NAVIGATE => {
            let Some(state) = progress_dialog_state_mut(hwnd) else {
                return 0;
            };
            if !state.action_hwnd.is_null() {
                SetFocus(state.action_hwnd);
                speak_control_text(state.owner, state.action_hwnd);
            }
            0
        }
        WM_PROGRESS_NAVIGATE => {
            let Some(state) = progress_dialog_state_mut(hwnd) else {
                return 0;
            };
            if state.lines.is_empty() {
                return 0;
            }
            if wparam == 0 {
                state.current_line = state.current_line.saturating_sub(1);
            } else {
                state.current_line =
                    (state.current_line + 1).min(state.lines.len().saturating_sub(1));
            }
            if !state.info_hwnd.is_null() {
                set_progress_text(state.info_hwnd, &state.lines, state.current_line);
            }
            speak_progress_line(state.owner, &state.lines, state.current_line);
            0
        }
        WM_KEYDOWN if matches!(wparam as u32, 0x26 | 0x28) => {
            let Some(state) = progress_dialog_state_mut(hwnd) else {
                return 0;
            };
            if state.lines.is_empty() {
                return 0;
            }
            if wparam as u32 == 0x26 {
                state.current_line = state.current_line.saturating_sub(1);
            } else {
                state.current_line =
                    (state.current_line + 1).min(state.lines.len().saturating_sub(1));
            }
            if !state.info_hwnd.is_null() {
                set_progress_text(state.info_hwnd, &state.lines, state.current_line);
            }
            speak_progress_line(state.owner, &state.lines, state.current_line);
            0
        }
        WM_COMMAND => {
            let id = loword(wparam as u32) as i32;
            if id == ID_DIALOG_CANCEL || id == ID_DIALOG_OK {
                if let Some(state) = progress_dialog_state_mut(hwnd) {
                    if state.running {
                        state.cancel_flag.store(true, Ordering::Relaxed);
                        SetWindowTextW(state.action_hwnd, wide("Anulowanie...").as_ptr());
                        EnableWindow(state.action_hwnd, 0);
                    } else {
                        DestroyWindow(hwnd);
                    }
                    return 0;
                }
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_CLOSE => {
            if let Some(state) = progress_dialog_state_mut(hwnd) {
                if state.running {
                    state.cancel_flag.store(true, Ordering::Relaxed);
                    SetWindowTextW(state.action_hwnd, wide("Anulowanie...").as_ptr());
                    EnableWindow(state.action_hwnd, 0);
                    return 0;
                }
            }
            DestroyWindow(hwnd);
            0
        }
        WM_PROGRESS_EVENT => {
            let Some(state) = progress_dialog_state_mut(hwnd) else {
                return 0;
            };
            while let Ok(event) = state.receiver.try_recv() {
                match event {
                    ProgressEvent::Update(lines) => {
                        state.lines = lines;
                        state.current_line =
                            state.current_line.min(state.lines.len().saturating_sub(1));
                        if !state.info_hwnd.is_null() {
                            set_progress_text(state.info_hwnd, &state.lines, state.current_line);
                        }
                    }
                    ProgressEvent::Conflict { destination } => {
                        let nvda = app_state_mut(state.owner).map(|app| &app.nvda);
                        let result = show_choice_prompt(
                            hwnd,
                            "Konflikt nazw",
                            vec![
                                "Element docelowy już istnieje.".to_string(),
                                destination.display().to_string(),
                                "Wybierz zastąp, pomiń albo anuluj operację.".to_string(),
                            ],
                            vec![
                                DialogButton {
                                    id: ID_DIALOG_REPLACE,
                                    label: "Zastąp",
                                    is_default: true,
                                },
                                DialogButton {
                                    id: ID_DIALOG_SKIP,
                                    label: "Pomiń",
                                    is_default: false,
                                },
                                DialogButton {
                                    id: ID_DIALOG_CANCEL,
                                    label: "Anuluj",
                                    is_default: false,
                                },
                            ],
                            nvda,
                        );
                        let choice = match result {
                            ID_DIALOG_REPLACE => ConflictChoice::Replace,
                            ID_DIALOG_SKIP => ConflictChoice::Skip,
                            _ => ConflictChoice::Cancel,
                        };
                        let _ = state.conflict_response_sender.send(choice);
                        SetFocus(state.action_hwnd);
                    }
                    ProgressEvent::Finished {
                        lines,
                        speech,
                        outcome,
                    } => {
                        let should_auto_close = state.auto_close_on_success
                            && matches!(outcome, WorkerOutcome::Success);
                        let should_auto_close_retryable = state.auto_close_on_retryable_error
                            && matches!(&outcome, WorkerOutcome::Error(message) if is_sftp_retry_message(message));
                        state.lines = lines;
                        state.running = false;
                        state.current_line =
                            state.current_line.min(state.lines.len().saturating_sub(1));
                        if let Ok(mut shared_outcome) = state.shared_outcome.lock() {
                            *shared_outcome = outcome;
                        }
                        if !state.info_hwnd.is_null() {
                            set_progress_text(state.info_hwnd, &state.lines, state.current_line);
                        }
                        if !speech.is_empty() && !should_auto_close_retryable {
                            if let Some(app) = app_state_mut(state.owner) {
                                app.nvda.speak_non_interrupting(&speech);
                            }
                        }
                        if should_auto_close || should_auto_close_retryable {
                            state.done = true;
                            DestroyWindow(hwnd);
                            return 0;
                        }
                        SetWindowTextW(state.action_hwnd, wide("Zamknij").as_ptr());
                        EnableWindow(state.action_hwnd, 1);
                        SetFocus(state.action_hwnd);
                        speak_control_text(state.owner, state.action_hwnd);
                    }
                }
            }
            0
        }
        WM_CTLCOLORSTATIC | WM_CTLCOLOREDIT | WM_CTLCOLORLISTBOX => {
            let hdc = wparam as HDC;
            SetTextColor(hdc, YELLOW);
            SetBkColor(hdc, BLACK);
            if msg == WM_CTLCOLORSTATIC {
                SetBkMode(hdc, TRANSPARENT as i32);
            }
            GetStockObject(4) as LRESULT
        }
        WM_DESTROY => {
            if let Some(state) = progress_dialog_state_mut(hwnd) {
                state.done = true;
            }
            0
        }
        WM_NCDESTROY => {
            let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut ProgressDialogState;
            if !ptr.is_null() {
                let _ = Box::from_raw(ptr);
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

unsafe extern "system" fn operation_prompt_dialog_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_NCCREATE => {
            let create =
                lparam as *const windows_sys::Win32::UI::WindowsAndMessaging::CREATESTRUCTW;
            let ptr = (*create).lpCreateParams as *mut OperationPromptState;
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, ptr as isize);
            1
        }
        WM_CREATE => {
            let Some(state) = operation_prompt_state_mut(hwnd) else {
                return -1;
            };

            let font = GetStockObject(17) as HFONT;
            let list_height = dialog_list_height(state.lines.len());
            let list_hwnd = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                wide("LISTBOX").as_ptr(),
                wide("").as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_VSCROLL | LBS_NOTIFY as u32,
                12,
                12,
                420,
                list_height,
                hwnd,
                ID_DIALOG_INFO as HMENU,
                GetModuleHandleW(null()),
                null_mut(),
            );
            SendMessageW(list_hwnd, 0x0030, font as WPARAM, 1);
            populate_listbox_lines(list_hwnd, &state.lines);
            let default_id = state
                .buttons
                .iter()
                .find(|button| button.is_default)
                .map(|button| button.id)
                .unwrap_or(ID_DIALOG_OK);
            let cancel_id = state
                .buttons
                .iter()
                .find(|button| button.id == ID_DIALOG_CANCEL)
                .map(|button| button.id)
                .unwrap_or(default_id);
            SetWindowSubclass(
                list_hwnd,
                Some(dialog_listbox_subclass_proc),
                3,
                pack_dialog_button_ids(default_id, cancel_id),
            );

            let button_width = 96;
            let button_gap = 12;
            let button_count = state.buttons.len() as i32;
            let total_width =
                button_count * button_width + (button_count.saturating_sub(1) * button_gap);
            let button_left = 12 + (420 - total_width).max(0);
            let button_top = 24 + list_height;
            let mut button_hwnds = Vec::with_capacity(state.buttons.len());

            for (index, button) in state.buttons.iter().enumerate() {
                let style = if button.is_default {
                    BS_DEFPUSHBUTTON as u32
                } else {
                    BS_PUSHBUTTON as u32
                };
                let control = CreateWindowExW(
                    0,
                    wide("BUTTON").as_ptr(),
                    wide(button.label).as_ptr(),
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | style,
                    button_left + (index as i32 * (button_width + button_gap)),
                    button_top,
                    button_width,
                    28,
                    hwnd,
                    button.id as HMENU,
                    GetModuleHandleW(null()),
                    null_mut(),
                );
                SendMessageW(control, 0x0030, font as WPARAM, 1);
                SetWindowSubclass(control, Some(dialog_button_subclass_proc), 6 + index, 0);
                button_hwnds.push(control);
            }
            state.info_hwnd = list_hwnd;
            state.button_hwnds = button_hwnds;
            SetFocus(list_hwnd);
            speak_dialog_list_selection(state.owner, list_hwnd, &state.lines);
            0
        }
        WM_DIALOG_NAVIGATE => {
            let Some(state) = operation_prompt_state_mut(hwnd) else {
                return 0;
            };
            let mut order = Vec::with_capacity(1 + state.button_hwnds.len());
            order.push(state.info_hwnd);
            order.extend(state.button_hwnds.iter().copied());
            focus_in_order(&order, lparam as HWND, wparam != 0);
            0
        }
        WM_COMMAND => {
            let id = loword(wparam as u32) as i32;
            let code = hiword(wparam as u32);
            if id == ID_DIALOG_INFO && code == LBN_SELCHANGE as u16 {
                if let Some(state) = operation_prompt_state_mut(hwnd) {
                    speak_dialog_list_selection(state.owner, lparam as HWND, &state.lines);
                    return 0;
                }
            }
            if let Some(state) = operation_prompt_state_mut(hwnd) {
                if state.buttons.iter().any(|button| button.id == id) {
                    state.result = id;
                    DestroyWindow(hwnd);
                    return 0;
                }
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_CLOSE => {
            if let Some(state) = operation_prompt_state_mut(hwnd) {
                state.result = ID_DIALOG_CANCEL;
            }
            DestroyWindow(hwnd);
            0
        }
        WM_CTLCOLORSTATIC | WM_CTLCOLOREDIT | WM_CTLCOLORLISTBOX => {
            let hdc = wparam as HDC;
            SetTextColor(hdc, YELLOW);
            SetBkColor(hdc, BLACK);
            if msg == WM_CTLCOLORSTATIC {
                SetBkMode(hdc, TRANSPARENT as i32);
            }
            GetStockObject(4) as LRESULT
        }
        WM_DESTROY => {
            if let Some(state) = operation_prompt_state_mut(hwnd) {
                state.done = true;
            }
            0
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

unsafe fn show_input_dialog(
    owner: HWND,
    title: &str,
    prompt: &str,
    initial: &str,
    nvda: Option<&NvdaController>,
) -> Option<String> {
    let mut state = Box::new(InputDialogState {
        owner,
        prompt: prompt.to_string(),
        prompt_lines: dialog_lines(prompt),
        initial: initial.to_string(),
        result: None,
        done: false,
        accepted: false,
        prompt_hwnd: null_mut(),
        edit_hwnd: null_mut(),
        ok_hwnd: null_mut(),
        cancel_hwnd: null_mut(),
    });

    let ptr = &mut *state as *mut InputDialogState;
    EnableWindow(owner, 0);
    if let Some(nvda) = nvda {
        nvda.speak_non_interrupting(&format!("{title}, {prompt}"));
    }
    let prompt_height = dialog_list_height(state.prompt_lines.len());
    let hwnd = CreateWindowExW(
        WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT,
        wide(INPUT_DIALOG_CLASS).as_ptr(),
        wide(title).as_ptr(),
        WS_CAPTION | WS_SYSMENU | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        400,
        128 + prompt_height,
        owner,
        null_mut(),
        GetModuleHandleW(null()),
        ptr as _,
    );

    if hwnd.is_null() {
        EnableWindow(owner, 1);
        return None;
    }

    let mut message: MSG = std::mem::zeroed();
    while !state.done && GetMessageW(&mut message, null_mut(), 0, 0) > 0 {
        if message.message == WM_KEYDOWN {
            let focused = GetFocus();
            if message.wParam as u32 == 0x09 {
                let order = [
                    state.prompt_hwnd,
                    state.edit_hwnd,
                    state.ok_hwnd,
                    state.cancel_hwnd,
                ];
                focus_in_order(&order, focused, shift_pressed());
                continue;
            }
            if message.wParam as u32 == 0x0D {
                if focused == state.edit_hwnd {
                    SendMessageW(
                        hwnd,
                        WM_COMMAND,
                        ID_DIALOG_OK as WPARAM,
                        state.ok_hwnd as LPARAM,
                    );
                    continue;
                }
                if focused == state.ok_hwnd {
                    SendMessageW(
                        hwnd,
                        WM_COMMAND,
                        ID_DIALOG_OK as WPARAM,
                        state.ok_hwnd as LPARAM,
                    );
                    continue;
                }
                if focused == state.cancel_hwnd {
                    SendMessageW(
                        hwnd,
                        WM_COMMAND,
                        ID_DIALOG_CANCEL as WPARAM,
                        state.cancel_hwnd as LPARAM,
                    );
                    continue;
                }
            }
            if message.wParam as u32 == 0x1B {
                SendMessageW(
                    hwnd,
                    WM_COMMAND,
                    ID_DIALOG_CANCEL as WPARAM,
                    state.cancel_hwnd as LPARAM,
                );
                continue;
            }
        }
        if IsDialogMessageW(hwnd, &mut message) == 0 {
            TranslateMessage(&message);
            DispatchMessageW(&message);
        }
    }

    EnableWindow(owner, 1);
    if let Some(app) = app_state_mut(owner) {
        SetFocus(app.panels[app.active_panel].list_hwnd);
    } else {
        SetFocus(owner);
    }
    if state.accepted {
        state.result.take()
    } else {
        None
    }
}

unsafe fn show_search_dialog(owner: HWND, nvda: Option<&NvdaController>) -> Option<(String, bool)> {
    let mut state = Box::new(SearchDialogState {
        owner,
        result: None,
        done: false,
        accepted: false,
        prompt_lines: vec![
            "Wpisz nazwę lub wyrażenie regularne do wyszukania.".to_string(),
            "Użyj przycisków Szukaj w katalogu albo Szukaj rekurencyjnie.".to_string(),
        ],
        prompt_hwnd: null_mut(),
        edit_hwnd: null_mut(),
        local_hwnd: null_mut(),
        recursive_hwnd: null_mut(),
        cancel_hwnd: null_mut(),
    });

    let ptr = &mut *state as *mut SearchDialogState;
    EnableWindow(owner, 0);
    if let Some(nvda) = nvda {
        nvda.speak_non_interrupting("Wyszukiwanie, pole edycji");
    }
    let prompt_height = dialog_list_height(state.prompt_lines.len());
    let hwnd = CreateWindowExW(
        WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT,
        wide(SEARCH_DIALOG_CLASS).as_ptr(),
        wide("Wyszukiwanie").as_ptr(),
        WS_CAPTION | WS_SYSMENU | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        480,
        128 + prompt_height,
        owner,
        null_mut(),
        GetModuleHandleW(null()),
        ptr as _,
    );

    if hwnd.is_null() {
        EnableWindow(owner, 1);
        return None;
    }

    let mut message: MSG = std::mem::zeroed();
    while !state.done && GetMessageW(&mut message, null_mut(), 0, 0) > 0 {
        if message.message == WM_KEYDOWN {
            let focused = GetFocus();
            if message.wParam as u32 == 0x09 {
                let order = [
                    state.prompt_hwnd,
                    state.edit_hwnd,
                    state.local_hwnd,
                    state.recursive_hwnd,
                    state.cancel_hwnd,
                ];
                focus_in_order(&order, focused, shift_pressed());
                continue;
            }
            if message.wParam as u32 == 0x0D {
                if focused == state.edit_hwnd || focused == state.local_hwnd {
                    SendMessageW(
                        hwnd,
                        WM_COMMAND,
                        ID_DIALOG_SEARCH_LOCAL as WPARAM,
                        state.local_hwnd as LPARAM,
                    );
                    continue;
                }
                if focused == state.recursive_hwnd {
                    SendMessageW(
                        hwnd,
                        WM_COMMAND,
                        ID_DIALOG_SEARCH_RECURSIVE as WPARAM,
                        state.recursive_hwnd as LPARAM,
                    );
                    continue;
                }
                if focused == state.cancel_hwnd {
                    SendMessageW(
                        hwnd,
                        WM_COMMAND,
                        ID_DIALOG_CANCEL as WPARAM,
                        state.cancel_hwnd as LPARAM,
                    );
                    continue;
                }
            }
            if message.wParam as u32 == 0x1B {
                SendMessageW(
                    hwnd,
                    WM_COMMAND,
                    ID_DIALOG_CANCEL as WPARAM,
                    state.cancel_hwnd as LPARAM,
                );
                continue;
            }
        }
        if IsDialogMessageW(hwnd, &mut message) == 0 {
            TranslateMessage(&message);
            DispatchMessageW(&message);
        }
    }

    EnableWindow(owner, 1);
    if let Some(app) = app_state_mut(owner) {
        SetFocus(app.panels[app.active_panel].list_hwnd);
    } else {
        SetFocus(owner);
    }
    if state.accepted {
        state.result.take()
    } else {
        None
    }
}

unsafe fn show_network_connection_dialog(
    owner: HWND,
    initial: Option<NetworkResource>,
    nvda: Option<&NvdaController>,
) -> Option<NetworkResource> {
    let initial = initial.unwrap_or_default();
    let mut state = Box::new(NetworkDialogState {
        owner,
        initial,
        result: None,
        done: false,
        accepted: false,
        prompt_lines: vec![
            "Wybierz typ serwera i uzupełnij dane połączenia.".to_string(),
            "Dane logowania są opcjonalne, jeśli serwer ich nie wymaga.".to_string(),
            "Klucz SSH dotyczy połączeń SFTP.".to_string(),
        ],
        info_hwnd: null_mut(),
        protocol_hwnd: null_mut(),
        host_hwnd: null_mut(),
        username_hwnd: null_mut(),
        password_hwnd: null_mut(),
        ssh_key_hwnd: null_mut(),
        ssh_key_browse_hwnd: null_mut(),
        directory_hwnd: null_mut(),
        display_name_hwnd: null_mut(),
        anonymous_hwnd: null_mut(),
        ok_hwnd: null_mut(),
        cancel_hwnd: null_mut(),
    });

    let ptr = &mut *state as *mut NetworkDialogState;
    EnableWindow(owner, 0);
    if let Some(nvda) = nvda {
        nvda.speak_non_interrupting("Dodaj połączenie sieciowe");
    }
    let hwnd = CreateWindowExW(
        WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT,
        wide(NETWORK_DIALOG_CLASS).as_ptr(),
        wide("Dodaj połączenie sieciowe").as_ptr(),
        WS_CAPTION | WS_SYSMENU | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        560,
        420,
        owner,
        null_mut(),
        GetModuleHandleW(null()),
        ptr as _,
    );

    if hwnd.is_null() {
        EnableWindow(owner, 1);
        return None;
    }

    let mut message: MSG = std::mem::zeroed();
    while !state.done && GetMessageW(&mut message, null_mut(), 0, 0) > 0 {
        if message.message == WM_KEYDOWN {
            let focused = GetFocus();
            if message.wParam as u32 == 0x09 {
                let mut order = Vec::with_capacity(12);
                order.push(state.info_hwnd);
                order.push(state.protocol_hwnd);
                order.push(state.host_hwnd);
                order.push(state.username_hwnd);
                order.push(state.password_hwnd);
                order.push(state.anonymous_hwnd);
                order.push(state.ssh_key_hwnd);
                order.push(state.ssh_key_browse_hwnd);
                order.push(state.directory_hwnd);
                order.push(state.display_name_hwnd);
                order.push(state.ok_hwnd);
                order.push(state.cancel_hwnd);
                focus_in_order(&order, focused, shift_pressed());
                continue;
            }
            if message.wParam as u32 == 0x1B {
                SendMessageW(
                    hwnd,
                    WM_COMMAND,
                    ID_DIALOG_CANCEL as WPARAM,
                    state.cancel_hwnd as LPARAM,
                );
                continue;
            }
        }
        if IsDialogMessageW(hwnd, &mut message) == 0 {
            TranslateMessage(&message);
            DispatchMessageW(&message);
        }
    }

    EnableWindow(owner, 1);
    if let Some(app) = app_state_mut(owner) {
        SetFocus(app.panels[app.active_panel].list_hwnd);
    } else {
        SetFocus(owner);
    }
    if state.accepted {
        state.result.take()
    } else {
        None
    }
}

unsafe fn show_archive_create_dialog(
    owner: HWND,
    initial_name: &str,
    nvda: Option<&NvdaController>,
) -> Option<ArchiveCreateOptions> {
    let mut state = Box::new(ArchiveCreateDialogState {
        owner,
        initial_name: initial_name.to_string(),
        result: None,
        done: false,
        accepted: false,
        prompt_lines: vec![
            "Skonfiguruj archiwum lub obraz dysku.".to_string(),
            "Szyfrowanie jest dostępne dla 7z i zip.".to_string(),
            "Rozmiar części podaj jak w 7-Zip, na przykład 100m albo 1g.".to_string(),
        ],
        info_hwnd: null_mut(),
        format_hwnd: null_mut(),
        name_hwnd: null_mut(),
        level_hwnd: null_mut(),
        encrypted_hwnd: null_mut(),
        encryption_hwnd: null_mut(),
        password_hwnd: null_mut(),
        volume_hwnd: null_mut(),
        ok_hwnd: null_mut(),
        cancel_hwnd: null_mut(),
    });

    let ptr = &mut *state as *mut ArchiveCreateDialogState;
    EnableWindow(owner, 0);
    if let Some(nvda) = nvda {
        nvda.speak_non_interrupting("Utwórz archiwum lub obraz dysku");
    }
    let prompt_height = dialog_list_height(state.prompt_lines.len());
    let hwnd = CreateWindowExW(
        WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT,
        wide(ARCHIVE_CREATE_DIALOG_CLASS).as_ptr(),
        wide("Utwórz archiwum lub obraz dysku").as_ptr(),
        WS_CAPTION | WS_SYSMENU | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        560,
        350 + prompt_height,
        owner,
        null_mut(),
        GetModuleHandleW(null()),
        ptr as _,
    );
    if hwnd.is_null() {
        EnableWindow(owner, 1);
        return None;
    }

    let mut message: MSG = std::mem::zeroed();
    while !state.done && GetMessageW(&mut message, null_mut(), 0, 0) > 0 {
        if message.message == WM_KEYDOWN {
            let focused = GetFocus();
            if message.wParam as u32 == 0x09 {
                let order = [
                    state.info_hwnd,
                    state.format_hwnd,
                    state.name_hwnd,
                    state.level_hwnd,
                    state.encrypted_hwnd,
                    state.encryption_hwnd,
                    state.password_hwnd,
                    state.volume_hwnd,
                    state.ok_hwnd,
                    state.cancel_hwnd,
                ];
                focus_in_order(&order, focused, shift_pressed());
                continue;
            }
            if message.wParam as u32 == 0x0D {
                if focused == state.name_hwnd
                    || focused == state.password_hwnd
                    || focused == state.volume_hwnd
                    || focused == state.ok_hwnd
                {
                    SendMessageW(
                        hwnd,
                        WM_COMMAND,
                        ID_DIALOG_OK as WPARAM,
                        state.ok_hwnd as LPARAM,
                    );
                    continue;
                }
                if focused == state.cancel_hwnd {
                    SendMessageW(
                        hwnd,
                        WM_COMMAND,
                        ID_DIALOG_CANCEL as WPARAM,
                        state.cancel_hwnd as LPARAM,
                    );
                    continue;
                }
            }
            if message.wParam as u32 == 0x1B {
                SendMessageW(
                    hwnd,
                    WM_COMMAND,
                    ID_DIALOG_CANCEL as WPARAM,
                    state.cancel_hwnd as LPARAM,
                );
                continue;
            }
        }
        if IsDialogMessageW(hwnd, &mut message) == 0 {
            TranslateMessage(&message);
            DispatchMessageW(&message);
        }
    }

    EnableWindow(owner, 1);
    if let Some(app) = app_state_mut(owner) {
        SetFocus(app.panels[app.active_panel].list_hwnd);
    } else {
        SetFocus(owner);
    }
    if state.accepted {
        state.result.take()
    } else {
        None
    }
}

unsafe fn show_discovery_dialog(
    owner: HWND,
    servers: Vec<DiscoveredServer>,
    nvda: Option<&NvdaController>,
) -> Option<DiscoveredServer> {
    let mut state = Box::new(DiscoveryDialogState {
        owner,
        result: None,
        done: false,
        servers,
        prompt_lines: vec![
            "Wybierz znaleziony serwer do dodania.".to_string(),
            "Lista zawiera host i wykryte protokoły.".to_string(),
        ],
        info_hwnd: null_mut(),
        list_hwnd: null_mut(),
        add_hwnd: null_mut(),
        cancel_hwnd: null_mut(),
    });

    let ptr = &mut *state as *mut DiscoveryDialogState;
    EnableWindow(owner, 0);
    if let Some(nvda) = nvda {
        nvda.speak_non_interrupting("Lista znalezionych serwerów");
    }
    let prompt_height = dialog_list_height(state.prompt_lines.len());
    let hwnd = CreateWindowExW(
        WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT,
        wide(DISCOVERY_DIALOG_CLASS).as_ptr(),
        wide("Wyszukiwanie serwerów").as_ptr(),
        WS_CAPTION | WS_SYSMENU | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        500,
        320 + prompt_height,
        owner,
        null_mut(),
        GetModuleHandleW(null()),
        ptr as _,
    );

    if hwnd.is_null() {
        EnableWindow(owner, 1);
        return None;
    }

    let mut message: MSG = std::mem::zeroed();
    while !state.done && GetMessageW(&mut message, null_mut(), 0, 0) > 0 {
        if message.message == WM_KEYDOWN {
            let focused = GetFocus();
            if message.wParam as u32 == 0x09 {
                let order = [
                    state.info_hwnd,
                    state.list_hwnd,
                    state.add_hwnd,
                    state.cancel_hwnd,
                ];
                focus_in_order(&order, focused, shift_pressed());
                continue;
            }
            if message.wParam as u32 == 0x1B {
                SendMessageW(
                    hwnd,
                    WM_COMMAND,
                    ID_DIALOG_CANCEL as WPARAM,
                    state.cancel_hwnd as LPARAM,
                );
                continue;
            }
            if message.wParam as u32 == 0x0D && focused == state.list_hwnd {
                SendMessageW(
                    hwnd,
                    WM_COMMAND,
                    ID_NET_DISCOVERY_ADD as WPARAM,
                    state.add_hwnd as LPARAM,
                );
                continue;
            }
        }
        if IsDialogMessageW(hwnd, &mut message) == 0 {
            TranslateMessage(&message);
            DispatchMessageW(&message);
        }
    }

    EnableWindow(owner, 1);
    if let Some(app) = app_state_mut(owner) {
        SetFocus(app.panels[app.active_panel].list_hwnd);
    } else {
        SetFocus(owner);
    }
    state.result.take()
}

unsafe fn show_operation_prompt(
    owner: HWND,
    title: &str,
    lines: Vec<String>,
    nvda: Option<&NvdaController>,
) -> bool {
    show_choice_prompt(
        owner,
        title,
        lines,
        vec![
            DialogButton {
                id: ID_DIALOG_OK,
                label: "Wykonaj",
                is_default: true,
            },
            DialogButton {
                id: ID_DIALOG_CANCEL,
                label: "Anuluj",
                is_default: false,
            },
        ],
        nvda,
    ) == ID_DIALOG_OK
}

unsafe fn show_info_prompt(
    owner: HWND,
    title: &str,
    lines: Vec<String>,
    nvda: Option<&NvdaController>,
) {
    let _ = show_choice_prompt(
        owner,
        title,
        lines,
        vec![DialogButton {
            id: ID_DIALOG_OK,
            label: "OK",
            is_default: true,
        }],
        nvda,
    );
}

unsafe fn show_choice_prompt(
    owner: HWND,
    title: &str,
    lines: Vec<String>,
    buttons: Vec<DialogButton>,
    nvda: Option<&NvdaController>,
) -> i32 {
    let mut state = Box::new(OperationPromptState {
        owner,
        lines,
        done: false,
        result: ID_DIALOG_CANCEL,
        buttons,
        info_hwnd: null_mut(),
        button_hwnds: Vec::new(),
    });

    let ptr = &mut *state as *mut OperationPromptState;
    EnableWindow(owner, 0);
    if let Some(nvda) = nvda {
        let spoken = if state.lines.is_empty() {
            title.to_string()
        } else {
            format!("{title}, {}", state.lines.join(", "))
        };
        nvda.speak_non_interrupting(&spoken);
    }
    let height = 92 + dialog_list_height(state.lines.len());
    let hwnd = CreateWindowExW(
        WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT,
        wide(OPERATION_PROMPT_DIALOG_CLASS).as_ptr(),
        wide(title).as_ptr(),
        WS_CAPTION | WS_SYSMENU | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        460,
        height,
        owner,
        null_mut(),
        GetModuleHandleW(null()),
        ptr as _,
    );

    if hwnd.is_null() {
        EnableWindow(owner, 1);
        return ID_DIALOG_CANCEL;
    }

    let mut message: MSG = std::mem::zeroed();
    while !state.done && GetMessageW(&mut message, null_mut(), 0, 0) > 0 {
        if message.message == WM_KEYDOWN {
            let focused = GetFocus();
            if message.wParam as u32 == 0x09 {
                let mut order = Vec::with_capacity(1 + state.button_hwnds.len());
                order.push(state.info_hwnd);
                order.extend(state.button_hwnds.iter().copied());
                focus_in_order(&order, focused, shift_pressed());
                continue;
            }
            if message.wParam as u32 == 0x0D {
                if let Some((index, _button_hwnd)) = state
                    .button_hwnds
                    .iter()
                    .enumerate()
                    .find(|(_, button_hwnd)| **button_hwnd == focused)
                {
                    if let Some(button) = state.buttons.get(index) {
                        SendMessageW(hwnd, WM_COMMAND, button.id as WPARAM, focused as LPARAM);
                        continue;
                    }
                }
            }
        }
        if IsDialogMessageW(hwnd, &mut message) == 0 {
            TranslateMessage(&message);
            DispatchMessageW(&message);
        }
    }

    EnableWindow(owner, 1);
    if let Some(app) = app_state_mut(owner) {
        SetFocus(app.panels[app.active_panel].list_hwnd);
    } else {
        SetFocus(owner);
    }
    state.result
}

impl ProgressReporter {
    fn new(
        sender: Sender<ProgressEvent>,
        notify_hwnd: HWND,
        cancel_flag: Arc<AtomicBool>,
        action: &'static str,
        total: ItemCounts,
        conflict_response_receiver: Receiver<ConflictChoice>,
    ) -> Self {
        Self {
            sender,
            notify_hwnd,
            cancel_flag,
            action,
            total,
            processed: ItemCounts::default(),
            started_at: Instant::now(),
            last_sent: Instant::now() - Duration::from_secs(1),
            last_lines: Vec::new(),
            status: "Przygotowanie operacji.".to_string(),
            current: None,
            destination: None,
            conflict_response_receiver,
        }
    }

    fn update(&mut self, status: &str, current: Option<&Path>, destination: Option<&Path>) {
        self.status = status.to_string();
        self.current = current.map(Path::to_path_buf);
        self.destination = destination.map(Path::to_path_buf);
        self.send_update_if_due(false);
    }

    fn add_counts(
        &mut self,
        counts: ItemCounts,
        status: &str,
        current: Option<&Path>,
        destination: Option<&Path>,
    ) {
        self.processed.add(counts);
        self.update(status, current, destination);
    }

    fn set_processed(
        &mut self,
        counts: ItemCounts,
        status: &str,
        current: Option<&Path>,
        destination: Option<&Path>,
    ) {
        self.processed = counts;
        self.update(status, current, destination);
    }

    fn force_update(&mut self) {
        self.send_update_if_due(true);
    }

    fn is_canceled(&self) -> bool {
        self.cancel_flag.load(Ordering::Relaxed)
    }

    fn ask_conflict(&mut self, destination: &Path) -> io::Result<ConflictChoice> {
        let _ = self.sender.send(ProgressEvent::Conflict {
            destination: destination.to_path_buf(),
        });
        unsafe {
            PostMessageW(self.notify_hwnd, WM_PROGRESS_EVENT, 0, 0);
        }
        loop {
            if self.is_canceled() {
                return Ok(ConflictChoice::Cancel);
            }
            match self
                .conflict_response_receiver
                .recv_timeout(Duration::from_millis(100))
            {
                Ok(choice) => return Ok(choice),
                Err(mpsc::RecvTimeoutError::Timeout) => {}
                Err(mpsc::RecvTimeoutError::Disconnected) => return Ok(ConflictChoice::Cancel),
            }
        }
    }

    fn finish(&mut self, status: &str, speech: String, outcome: WorkerOutcome) {
        self.status = status.to_string();
        self.current = None;
        self.destination = None;
        let lines = self.current_lines();
        let _ = self.sender.send(ProgressEvent::Finished {
            lines,
            speech,
            outcome,
        });
        unsafe {
            PostMessageW(self.notify_hwnd, WM_PROGRESS_EVENT, 0, 0);
        }
    }

    fn send_update_if_due(&mut self, force: bool) {
        let lines = self.current_lines();
        if !force && lines == self.last_lines {
            return;
        }
        if force || self.last_sent.elapsed() >= Duration::from_millis(250) {
            let _ = self.sender.send(ProgressEvent::Update(lines.clone()));
            unsafe {
                PostMessageW(self.notify_hwnd, WM_PROGRESS_EVENT, 0, 0);
            }
            self.last_lines = lines;
            self.last_sent = Instant::now();
        }
    }

    fn current_lines(&self) -> Vec<String> {
        build_progress_lines(
            self.action,
            &self.status,
            self.current.as_deref(),
            self.destination.as_deref(),
            self.processed,
            self.total,
            self.started_at.elapsed(),
        )
    }
}

unsafe fn run_progress_dialog<F>(
    owner: HWND,
    title: &str,
    initial_lines: Vec<String>,
    nvda: Option<&NvdaController>,
    options: ProgressDialogOptions,
    action: F,
) -> io::Result<WorkerOutcome>
where
    F: FnOnce(HWND, Sender<ProgressEvent>, Arc<AtomicBool>, Receiver<ConflictChoice>)
        + Send
        + 'static,
{
    let (sender, receiver) = mpsc::channel();
    let (conflict_response_sender, conflict_response_receiver) = mpsc::channel();
    let cancel_flag = Arc::new(AtomicBool::new(false));
    let shared_outcome = Arc::new(Mutex::new(WorkerOutcome::Canceled));
    let state = Box::new(ProgressDialogState {
        owner,
        lines: initial_lines,
        done: false,
        running: true,
        auto_close_on_success: options.auto_close_on_success,
        auto_close_on_retryable_error: options.auto_close_on_retryable_error,
        cancel_flag: cancel_flag.clone(),
        receiver,
        conflict_response_sender,
        shared_outcome: shared_outcome.clone(),
        info_hwnd: null_mut(),
        action_hwnd: null_mut(),
        current_line: 0,
    });
    let height = 92 + dialog_list_height(state.lines.len());
    let first_line = state.lines.first().cloned();
    let ptr = Box::into_raw(state);
    EnableWindow(owner, 0);
    if let Some(nvda) = nvda {
        let spoken = if let Some(first_line) = first_line {
            format!("{title}, {first_line}")
        } else {
            title.to_string()
        };
        nvda.speak_non_interrupting(&spoken);
    }
    let hwnd = CreateWindowExW(
        WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT,
        wide(PROGRESS_DIALOG_CLASS).as_ptr(),
        wide(title).as_ptr(),
        WS_CAPTION | WS_SYSMENU | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        540,
        height,
        owner,
        null_mut(),
        GetModuleHandleW(null()),
        ptr as _,
    );
    if hwnd.is_null() {
        EnableWindow(owner, 1);
        let _ = Box::from_raw(ptr);
        return Err(io::Error::other("nie udało się utworzyć okna postępu"));
    }

    let worker_hwnd = hwnd as usize;
    let worker = thread::spawn(move || {
        action(
            worker_hwnd as HWND,
            sender,
            cancel_flag,
            conflict_response_receiver,
        );
    });

    let mut message: MSG = std::mem::zeroed();
    while progress_dialog_state_mut(hwnd)
        .map(|state| !state.done)
        .unwrap_or(false)
        && GetMessageW(&mut message, null_mut(), 0, 0) > 0
    {
        if message.message == WM_KEYDOWN {
            let focused = GetFocus();
            if message.wParam as u32 == 0x09 {
                if let Some(state) = progress_dialog_state_mut(hwnd) {
                    let order = [state.action_hwnd];
                    focus_in_order(&order, focused, shift_pressed());
                    continue;
                }
            }
            if message.wParam as u32 == 0x0D {
                if let Some(state) = progress_dialog_state_mut(hwnd) {
                    if focused == state.action_hwnd {
                        SendMessageW(
                            hwnd,
                            WM_COMMAND,
                            ID_DIALOG_CANCEL as WPARAM,
                            focused as LPARAM,
                        );
                        continue;
                    }
                }
            }
        }
        if IsDialogMessageW(hwnd, &mut message) == 0 {
            TranslateMessage(&message);
            DispatchMessageW(&message);
        }
    }

    let _ = worker.join();

    EnableWindow(owner, 1);
    if let Some(app) = app_state_mut(owner) {
        SetFocus(app.panels[app.active_panel].list_hwnd);
    } else {
        SetFocus(owner);
    }

    let outcome = shared_outcome
        .lock()
        .map(|outcome| outcome.clone())
        .unwrap_or(WorkerOutcome::Canceled);
    Ok(outcome)
}

unsafe fn run_simple_worker_dialog<F>(
    owner: HWND,
    title: &str,
    initial_lines: Vec<String>,
    nvda: Option<&NvdaController>,
    action: F,
    success_speech: &str,
) -> io::Result<WorkerOutcome>
where
    F: FnOnce() -> io::Result<()> + Send + 'static,
{
    let speech = success_speech.to_string();
    run_progress_dialog(
        owner,
        title,
        initial_lines,
        nvda,
        ProgressDialogOptions::default(),
        move |progress_hwnd, sender, cancel_flag, _conflict_receiver| {
            if cancel_flag.load(Ordering::Relaxed) {
                let _ = sender.send(ProgressEvent::Finished {
                    lines: vec!["Operacja anulowana.".to_string()],
                    speech: "anulowano".to_string(),
                    outcome: WorkerOutcome::Canceled,
                });
                unsafe {
                    PostMessageW(progress_hwnd, WM_PROGRESS_EVENT, 0, 0);
                }
                return;
            }
            match action() {
                Ok(()) => {
                    let _ = sender.send(ProgressEvent::Finished {
                        lines: vec![speech.clone()],
                        speech,
                        outcome: WorkerOutcome::Success,
                    });
                }
                Err(error) => {
                    let _ = sender.send(ProgressEvent::Finished {
                        lines: vec![format!("Błąd: {error}")],
                        speech: format!("operacja zakończona błędem: {error}"),
                        outcome: WorkerOutcome::Error(error.to_string()),
                    });
                }
            }
            unsafe {
                PostMessageW(progress_hwnd, WM_PROGRESS_EVENT, 0, 0);
            }
        },
    )
}

unsafe fn create_static(parent: HWND, id: i32) -> Result<HWND, String> {
    let hwnd = CreateWindowExW(
        0,
        wide("STATIC").as_ptr(),
        wide("").as_ptr(),
        WS_CHILD | WS_VISIBLE,
        0,
        0,
        0,
        0,
        parent,
        id as HMENU,
        GetModuleHandleW(null()),
        null_mut(),
    );
    if hwnd.is_null() {
        Err(last_error_message("nie udało się utworzyć etykiety"))
    } else {
        Ok(hwnd)
    }
}

unsafe fn create_listbox(parent: HWND, id: i32) -> Result<HWND, String> {
    let hwnd = CreateWindowExW(
        WS_EX_CLIENTEDGE,
        wide("LISTBOX").as_ptr(),
        wide("").as_ptr(),
        WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_TABSTOP | LBS_NOTIFY as u32,
        0,
        0,
        0,
        0,
        parent,
        id as HMENU,
        GetModuleHandleW(null()),
        null_mut(),
    );
    if hwnd.is_null() {
        Err(last_error_message("nie udało się utworzyć listy"))
    } else {
        Ok(hwnd)
    }
}

unsafe fn create_context_popup_menu() -> HMENU {
    let menu = CreatePopupMenu();
    AppendMenuW(
        menu,
        MF_STRING,
        IDM_SYSTEM_CONTEXT as usize,
        wide("Menu kontekstowe Windows").as_ptr(),
    );
    AppendMenuW(
        menu,
        MF_STRING,
        IDM_ADD_TO_FAVORITES as usize,
        wide("Dodaj do ulubionych").as_ptr(),
    );
    AppendMenuW(
        menu,
        MF_STRING,
        IDM_EDIT_NETWORK_CONNECTION as usize,
        wide("Edytuj połączenie sieciowe").as_ptr(),
    );
    AppendMenuW(
        menu,
        MF_STRING,
        IDM_REMOVE_NETWORK_CONNECTION as usize,
        wide("Usuń połączenie sieciowe").as_ptr(),
    );
    AppendMenuW(menu, MF_SEPARATOR, 0, null());
    AppendMenuW(
        menu,
        MF_STRING,
        IDM_RENAME as usize,
        wide("Zmień nazwę\tF2").as_ptr(),
    );
    AppendMenuW(
        menu,
        MF_STRING,
        IDM_COPY as usize,
        wide("Kopiuj\tF5").as_ptr(),
    );
    AppendMenuW(
        menu,
        MF_STRING,
        IDM_MOVE as usize,
        wide("Przenieś\tF6").as_ptr(),
    );
    AppendMenuW(
        menu,
        MF_STRING,
        IDM_NEW_FOLDER as usize,
        wide("Nowy katalog\tF7").as_ptr(),
    );
    AppendMenuW(
        menu,
        MF_STRING,
        IDM_DELETE as usize,
        wide("Usuń\tDelete").as_ptr(),
    );
    AppendMenuW(menu, MF_SEPARATOR, 0, null());
    AppendMenuW(
        menu,
        MF_STRING,
        IDM_EXTRACT_HERE as usize,
        wide("Wypakuj tutaj").as_ptr(),
    );
    AppendMenuW(
        menu,
        MF_STRING,
        IDM_EXTRACT_TO_FOLDER as usize,
        wide("Wypakuj do katalogu z nazwą archiwum").as_ptr(),
    );
    AppendMenuW(
        menu,
        MF_STRING,
        IDM_EXTRACT_TO_OTHER_PANEL as usize,
        wide("Wypakuj do drugiego panelu").as_ptr(),
    );
    AppendMenuW(
        menu,
        MF_STRING,
        IDM_CREATE_ARCHIVE as usize,
        wide("Utwórz archiwum lub obraz dysku").as_ptr(),
    );
    AppendMenuW(
        menu,
        MF_STRING,
        IDM_JOIN_SPLIT_ARCHIVE as usize,
        wide("Połącz podzielone archiwum").as_ptr(),
    );
    AppendMenuW(menu, MF_SEPARATOR, 0, null());
    AppendMenuW(
        menu,
        MF_STRING,
        IDM_CHECKSUM_CREATE as usize,
        wide("Utwórz sumę kontrolną SHA-256").as_ptr(),
    );
    AppendMenuW(
        menu,
        MF_STRING,
        IDM_CHECKSUM_VERIFY as usize,
        wide("Sprawdź sumę kontrolną SHA-256").as_ptr(),
    );
    AppendMenuW(menu, MF_SEPARATOR, 0, null());
    AppendMenuW(
        menu,
        MF_STRING,
        IDM_PROPERTIES as usize,
        wide("Właściwości").as_ptr(),
    );
    AppendMenuW(
        menu,
        MF_STRING,
        IDM_PERMISSIONS as usize,
        wide("Uprawnienia").as_ptr(),
    );
    AppendMenuW(menu, MF_SEPARATOR, 0, null());
    AppendMenuW(
        menu,
        MF_STRING,
        IDM_REFRESH as usize,
        wide("Odśwież\tF12").as_ptr(),
    );
    menu
}

#[repr(C)]
struct IUnknownVtbl {
    query_interface: unsafe extern "system" fn(
        this: *mut c_void,
        iid: *const GUID,
        interface: *mut *mut c_void,
    ) -> i32,
    add_ref: unsafe extern "system" fn(this: *mut c_void) -> u32,
    release: unsafe extern "system" fn(this: *mut c_void) -> u32,
}

#[repr(C)]
struct IShellFolderVtbl {
    parent: IUnknownVtbl,
    parse_display_name: usize,
    enum_objects: usize,
    bind_to_object: usize,
    bind_to_storage: usize,
    compare_ids: usize,
    create_view_object: usize,
    get_attributes_of: usize,
    get_ui_object_of: unsafe extern "system" fn(
        this: *mut c_void,
        hwnd_owner: HWND,
        cidl: u32,
        apidl: *const *const ITEMIDLIST,
        riid: *const GUID,
        reserved: *mut u32,
        object: *mut *mut c_void,
    ) -> i32,
    get_display_name_of: usize,
    set_name_of: usize,
}

#[repr(C)]
struct IContextMenuVtbl {
    parent: IUnknownVtbl,
    query_context_menu: unsafe extern "system" fn(
        this: *mut c_void,
        menu: HMENU,
        index_menu: u32,
        id_cmd_first: u32,
        id_cmd_last: u32,
        flags: u32,
    ) -> i32,
    invoke_command:
        unsafe extern "system" fn(this: *mut c_void, command: *const CMINVOKECOMMANDINFO) -> i32,
    get_command_string: usize,
}

#[repr(C)]
struct IShellLinkWVtbl {
    parent: IUnknownVtbl,
    get_path: unsafe extern "system" fn(
        this: *mut c_void,
        file: *mut u16,
        max_path: i32,
        find_data: *mut c_void,
        flags: u32,
    ) -> i32,
    get_id_list: usize,
    set_id_list: usize,
    get_description: usize,
    set_description: usize,
    get_working_directory: usize,
    set_working_directory: usize,
    get_arguments: usize,
    set_arguments: usize,
    get_hotkey: usize,
    set_hotkey: usize,
    get_show_cmd: usize,
    set_show_cmd: usize,
    get_icon_location: usize,
    set_icon_location: usize,
    set_relative_path: usize,
    resolve: usize,
    set_path: usize,
}

#[repr(C)]
struct IPersistFileVtbl {
    parent: IUnknownVtbl,
    get_class_id: usize,
    is_dirty: usize,
    load: unsafe extern "system" fn(this: *mut c_void, file_name: *const u16, mode: u32) -> i32,
    save: usize,
    save_completed: usize,
    get_cur_file: usize,
}

const IID_ISHELLFOLDER: GUID = GUID {
    data1: 0x000214E6,
    data2: 0x0000,
    data3: 0x0000,
    data4: [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46],
};

const IID_ICONTEXTMENU: GUID = GUID {
    data1: 0x000214E4,
    data2: 0x0000,
    data3: 0x0000,
    data4: [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46],
};

const IID_ISHELLLINKW: GUID = GUID {
    data1: 0x000214F9,
    data2: 0x0000,
    data3: 0x0000,
    data4: [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46],
};

const IID_IPERSISTFILE: GUID = GUID {
    data1: 0x0000010B,
    data2: 0x0000,
    data3: 0x0000,
    data4: [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46],
};

struct ComApartment {
    initialized: bool,
}

impl ComApartment {
    unsafe fn new() -> io::Result<Self> {
        let result = CoInitializeEx(
            null(),
            (COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE) as u32,
        );
        if result < 0 {
            Err(io::Error::other(format!(
                "nie udało się zainicjować COM: 0x{result:08X}"
            )))
        } else {
            Ok(Self { initialized: true })
        }
    }
}

impl Drop for ComApartment {
    fn drop(&mut self) {
        if self.initialized {
            unsafe {
                CoUninitialize();
            }
        }
    }
}

unsafe fn release_com(ptr: *mut c_void) {
    if ptr.is_null() {
        return;
    }
    let vtable = *(ptr as *mut *const IUnknownVtbl);
    ((*vtable).release)(ptr);
}

struct ComInterfacePtr(*mut c_void);

impl ComInterfacePtr {
    fn new(ptr: *mut c_void) -> io::Result<Self> {
        if ptr.is_null() {
            Err(io::Error::other("COM zwrócił pusty wskaźnik"))
        } else {
            Ok(Self(ptr))
        }
    }

    fn as_ptr(&self) -> *mut c_void {
        self.0
    }
}

impl Drop for ComInterfacePtr {
    fn drop(&mut self) {
        unsafe {
            release_com(self.0);
        }
    }
}

fn infer_file_type(name: &str) -> String {
    Path::new(name)
        .extension()
        .map(|extension| extension.to_string_lossy().to_lowercase())
        .filter(|extension| !extension.is_empty())
        .unwrap_or_else(|| "plik".to_string())
}

fn is_windows_shortcut_name(name: &str) -> bool {
    Path::new(name)
        .extension()
        .and_then(|extension| extension.to_str())
        .map(|extension| extension.eq_ignore_ascii_case("lnk"))
        .unwrap_or(false)
}

fn is_windows_shortcut_path(path: &Path) -> bool {
    path.extension()
        .and_then(|extension| extension.to_str())
        .map(|extension| extension.eq_ignore_ascii_case("lnk"))
        .unwrap_or(false)
}

fn resolve_local_symlink_path(link_path: &Path, target: &Path) -> PathBuf {
    if target.is_absolute() {
        target.to_path_buf()
    } else {
        link_path
            .parent()
            .unwrap_or_else(|| Path::new(""))
            .join(target)
    }
}

fn normalize_remote_path_text(path: &str) -> PathBuf {
    let absolute = path.starts_with('/');
    let mut parts = Vec::new();
    for part in path.replace('\\', "/").split('/') {
        match part {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            value => parts.push(value.to_string()),
        }
    }

    let mut normalized = if absolute {
        String::from("/")
    } else {
        String::new()
    };
    normalized.push_str(&parts.join("/"));
    if normalized.is_empty() {
        normalized.push('.');
    }
    PathBuf::from(normalized)
}

fn resolve_remote_symlink_path(link_path: &Path, target: &Path) -> PathBuf {
    let target_text = remote_shell_path(target);
    if target_text.starts_with('/') {
        return normalize_remote_path_text(&target_text);
    }

    let parent = link_path
        .parent()
        .map(remote_shell_path)
        .unwrap_or_else(|| "/".to_string());
    let combined = if parent == "/" || parent.is_empty() {
        format!("/{target_text}")
    } else {
        format!("{}/{}", parent.trim_end_matches('/'), target_text)
    };
    normalize_remote_path_text(&combined)
}

fn resolve_windows_shortcut(path: &Path) -> io::Result<PathBuf> {
    unsafe {
        let _com = ComApartment::new()?;
        let mut shell_link_ptr = null_mut();
        let create_result = CoCreateInstance(
            &ShellLink,
            null_mut(),
            CLSCTX_INPROC_SERVER,
            &IID_ISHELLLINKW,
            &mut shell_link_ptr,
        );
        if create_result < 0 {
            return Err(io::Error::other(format!(
                "nie udało się utworzyć ShellLink: 0x{create_result:08X}"
            )));
        }
        let shell_link = ComInterfacePtr::new(shell_link_ptr)?;

        let shell_link_vtable = *(shell_link.as_ptr() as *mut *const IUnknownVtbl);
        let mut persist_file_ptr = null_mut();
        let query_result = ((*shell_link_vtable).query_interface)(
            shell_link.as_ptr(),
            &IID_IPERSISTFILE,
            &mut persist_file_ptr,
        );
        if query_result < 0 {
            return Err(io::Error::other(format!(
                "nie udało się odczytać IPersistFile: 0x{query_result:08X}"
            )));
        }
        let persist_file = ComInterfacePtr::new(persist_file_ptr)?;

        let shortcut_path = path_wide(path);
        let persist_vtable = *(persist_file.as_ptr() as *mut *const IPersistFileVtbl);
        let load_result =
            ((*persist_vtable).load)(persist_file.as_ptr(), shortcut_path.as_ptr(), STGM_READ);
        if load_result < 0 {
            return Err(io::Error::other(format!(
                "nie udało się wczytać skrótu: 0x{load_result:08X}"
            )));
        }

        let mut target = vec![0u16; 32768];
        let shell_link_vtable = *(shell_link.as_ptr() as *mut *const IShellLinkWVtbl);
        let get_result = ((*shell_link_vtable).get_path)(
            shell_link.as_ptr(),
            target.as_mut_ptr(),
            target.len() as i32,
            null_mut(),
            SLGP_UNCPRIORITY as u32,
        );
        if get_result < 0 {
            return Err(io::Error::other(format!(
                "nie udało się odczytać celu skrótu: 0x{get_result:08X}"
            )));
        }

        let len = target
            .iter()
            .position(|ch| *ch == 0)
            .unwrap_or(target.len());
        if len == 0 {
            return Err(io::Error::other("skrót nie zawiera celu pliku"));
        }
        Ok(PathBuf::from(String::from_utf16_lossy(&target[..len])))
    }
}

fn local_path_is_directory(path: &Path) -> bool {
    unsafe {
        let attrs = GetFileAttributesW(path_wide(path).as_ptr());
        if attrs != INVALID_FILE_ATTRIBUTES {
            return attrs & FILE_ATTRIBUTE_DIRECTORY != 0;
        }
    }
    path.is_dir()
}

fn remote_client_path_is_directory(client: &mut RemoteClient, path: &Path) -> io::Result<bool> {
    match client.stat(path) {
        Ok(entry) if entry.is_dir() => return Ok(true),
        Ok(_) => {}
        Err(error) if is_not_found_error(&error) => return Ok(false),
        Err(error) => return Err(error),
    }

    match client.list_dir(path) {
        Ok(_) => Ok(true),
        Err(error)
            if is_not_found_error(&error)
                || error.to_string().to_lowercase().contains("not a directory")
                || error.to_string().to_lowercase().contains("notdir") =>
        {
            Ok(false)
        }
        Err(_) => Ok(false),
    }
}

fn remote_path_is_directory(remote: &RemoteLocation, path: &Path) -> io::Result<bool> {
    if is_sftp_sudo_location(remote) {
        return stat_sftp_path_via_sudo(remote, path).map(|(is_dir, _)| is_dir);
    }
    let mut client = connect_remote_client(&remote.resource)?;
    let result = remote_client_path_is_directory(&mut client, path);
    client.disconnect();
    result
}

fn resolve_remote_link_target(
    remote: &RemoteLocation,
    initial_target: &Path,
) -> io::Result<(PathBuf, EntryKind)> {
    if is_sftp_sudo_location(remote) {
        let (is_dir, _) = stat_sftp_path_via_sudo(remote, initial_target)?;
        return Ok((
            initial_target.to_path_buf(),
            if is_dir {
                EntryKind::Directory
            } else {
                EntryKind::File
            },
        ));
    }

    let mut client = connect_remote_client(&remote.resource)?;
    let mut target = initial_target.to_path_buf();
    let result = (|| {
        for _ in 0..8 {
            let entry = client.stat(&target)?;
            if entry.metadata().is_symlink() {
                if let Some(next_target) = &entry.metadata().symlink {
                    target = resolve_remote_symlink_path(&target, next_target);
                    continue;
                }
            }
            let is_dir = entry.is_dir() || remote_client_path_is_directory(&mut client, &target)?;
            return Ok((
                target,
                if is_dir {
                    EntryKind::Directory
                } else {
                    EntryKind::File
                },
            ));
        }
        Err(io::Error::other(
            "zbyt wiele zagnieżdżonych linków symbolicznych",
        ))
    })();
    client.disconnect();
    result
}

fn compile_search_regex(pattern: &str) -> io::Result<Regex> {
    if let Ok(regex) = RegexBuilder::new(pattern).case_insensitive(true).build() {
        return Ok(regex);
    }

    let mut translated = String::new();
    for ch in pattern.chars() {
        match ch {
            '*' => translated.push_str(".*"),
            '?' => translated.push('.'),
            _ => translated.push_str(&regex::escape(&ch.to_string())),
        }
    }

    RegexBuilder::new(&translated)
        .case_insensitive(true)
        .build()
        .map_err(|error| io::Error::other(format!("nieprawidłowe wyrażenie lub wzorzec: {error}")))
}

fn polish_month(month: u16) -> &'static str {
    match month {
        1 => "stycznia",
        2 => "lutego",
        3 => "marca",
        4 => "kwietnia",
        5 => "maja",
        6 => "czerwca",
        7 => "lipca",
        8 => "sierpnia",
        9 => "września",
        10 => "października",
        11 => "listopada",
        12 => "grudnia",
        _ => "nieznanego miesiąca",
    }
}

fn filetime_from_u64(value: u64) -> FILETIME {
    FILETIME {
        dwLowDateTime: value as u32,
        dwHighDateTime: (value >> 32) as u32,
    }
}

fn format_filetime_label(value: u64) -> Option<String> {
    if value == 0 {
        return None;
    }

    unsafe {
        let filetime = filetime_from_u64(value);
        let mut local_filetime = FILETIME {
            dwLowDateTime: 0,
            dwHighDateTime: 0,
        };
        if FileTimeToLocalFileTime(&filetime, &mut local_filetime) == 0 {
            return None;
        }

        let mut local_time = SYSTEMTIME {
            wYear: 0,
            wMonth: 0,
            wDayOfWeek: 0,
            wDay: 0,
            wHour: 0,
            wMinute: 0,
            wSecond: 0,
            wMilliseconds: 0,
        };
        if FileTimeToSystemTime(&local_filetime, &mut local_time) == 0 {
            return None;
        }

        Some(format!(
            "{} {} {} roku {:02}:{:02}",
            local_time.wDay,
            polish_month(local_time.wMonth),
            local_time.wYear,
            local_time.wHour,
            local_time.wMinute
        ))
    }
}

fn format_system_time_label(value: SystemTime) -> Option<String> {
    let duration = value.duration_since(UNIX_EPOCH).ok()?;
    let windows_epoch_offset = 11_644_473_600u64;
    let total_100ns = duration
        .as_secs()
        .checked_add(windows_epoch_offset)?
        .checked_mul(10_000_000)?
        .checked_add((duration.subsec_nanos() / 100) as u64)?;
    format_filetime_label(total_100ns)
}

fn sanitize_file_component(value: &str) -> String {
    let mut output = String::with_capacity(value.len());
    for ch in value.chars() {
        if ch.is_ascii_alphanumeric() || matches!(ch, '-' | '_' | '.') {
            output.push(ch);
        } else {
            output.push('_');
        }
    }
    if output.is_empty() {
        "remote".to_string()
    } else {
        output
    }
}

unsafe fn open_properties_sheet(
    owner: HWND,
    path: &Path,
    page_name: Option<&str>,
) -> io::Result<()> {
    let target = path_wide(path);
    let page = page_name.map(wide);
    let page_ptr = page.as_ref().map_or(null(), |text| text.as_ptr());
    if SHObjectProperties(owner, SHOP_FILEPATH as u32, target.as_ptr(), page_ptr) == 0 {
        Err(io::Error::other(format!(
            "system nie otworzył właściwości dla {}",
            path.display()
        )))
    } else {
        Ok(())
    }
}

unsafe fn show_shell_context_menu(
    owner: HWND,
    anchor_hwnd: HWND,
    targets: &[PathBuf],
) -> io::Result<()> {
    if targets.is_empty() {
        return Err(io::Error::other("brak elementów"));
    }

    let _com = ComApartment::new()?;
    let mut parent_folder: *mut c_void = null_mut();
    let mut absolute_pidls = Vec::<*mut ITEMIDLIST>::with_capacity(targets.len());
    let mut child_pidls = Vec::<*const ITEMIDLIST>::with_capacity(targets.len());

    for target in targets {
        let wide_target = path_wide(target);
        let mut absolute_pidl: *mut ITEMIDLIST = null_mut();
        let parse_result = SHParseDisplayName(
            wide_target.as_ptr(),
            null_mut(),
            &mut absolute_pidl,
            0,
            null_mut(),
        );
        if parse_result < 0 || absolute_pidl.is_null() {
            for pidl in absolute_pidls {
                ILFree(pidl);
            }
            if !parent_folder.is_null() {
                release_com(parent_folder);
            }
            return Err(io::Error::other(format!(
                "system nie rozpoznał ścieżki {}",
                target.display()
            )));
        }

        let mut current_parent: *mut c_void = null_mut();
        let mut child_pidl: *mut ITEMIDLIST = null_mut();
        let bind_result = SHBindToParent(
            absolute_pidl,
            &IID_ISHELLFOLDER,
            &mut current_parent,
            &mut child_pidl,
        );
        if bind_result < 0 || current_parent.is_null() || child_pidl.is_null() {
            ILFree(absolute_pidl);
            for pidl in absolute_pidls {
                ILFree(pidl);
            }
            if !parent_folder.is_null() {
                release_com(parent_folder);
            }
            return Err(io::Error::other(format!(
                "nie udało się przygotować menu dla {}",
                target.display()
            )));
        }

        if parent_folder.is_null() {
            parent_folder = current_parent;
        } else {
            release_com(current_parent);
        }

        absolute_pidls.push(absolute_pidl);
        child_pidls.push(child_pidl as *const ITEMIDLIST);
    }

    let mut context_menu_ptr: *mut c_void = null_mut();
    let folder_vtable = *(parent_folder as *mut *const IShellFolderVtbl);
    let ui_result = ((*folder_vtable).get_ui_object_of)(
        parent_folder,
        owner,
        child_pidls.len() as u32,
        child_pidls.as_ptr(),
        &IID_ICONTEXTMENU,
        null_mut(),
        &mut context_menu_ptr,
    );

    if ui_result < 0 || context_menu_ptr.is_null() {
        for pidl in absolute_pidls {
            ILFree(pidl);
        }
        release_com(parent_folder);
        return Err(io::Error::other(
            "nie udało się otworzyć menu kontekstowego Windows",
        ));
    }

    let menu = CreatePopupMenu();
    let context_vtable = *(context_menu_ptr as *mut *const IContextMenuVtbl);
    let query_result = ((*context_vtable).query_context_menu)(
        context_menu_ptr,
        menu,
        0,
        1,
        0x7FFF,
        CMF_NORMAL | CMF_EXPLORE | CMF_CANRENAME,
    );
    if query_result < 0 {
        DestroyMenu(menu);
        release_com(context_menu_ptr);
        release_com(parent_folder);
        for pidl in absolute_pidls {
            ILFree(pidl);
        }
        return Err(io::Error::other("system nie zbudował menu kontekstowego"));
    }

    let mut rect: RECT = std::mem::zeroed();
    GetWindowRect(anchor_hwnd, &mut rect);
    let command = TrackPopupMenuEx(
        menu,
        TPM_RETURNCMD | TPM_TOPALIGN | TPM_LEFTALIGN | TPM_RIGHTBUTTON,
        rect.left + 24,
        rect.top + 24,
        owner,
        null(),
    ) as u32;

    if command != 0 {
        let invoke = CMINVOKECOMMANDINFO {
            cbSize: std::mem::size_of::<CMINVOKECOMMANDINFO>() as u32,
            fMask: 0,
            hwnd: owner,
            lpVerb: (command - 1) as usize as *const u8,
            lpParameters: null(),
            lpDirectory: null(),
            nShow: SW_SHOW,
            dwHotKey: 0,
            hIcon: null_mut(),
        };
        let invoke_result = ((*context_vtable).invoke_command)(context_menu_ptr, &invoke);
        if invoke_result < 0 {
            DestroyMenu(menu);
            release_com(context_menu_ptr);
            release_com(parent_folder);
            for pidl in absolute_pidls {
                ILFree(pidl);
            }
            return Err(io::Error::other(
                "nie udało się wykonać polecenia systemowego menu",
            ));
        }
    }

    DestroyMenu(menu);
    release_com(context_menu_ptr);
    release_com(parent_folder);
    for pidl in absolute_pidls {
        ILFree(pidl);
    }
    Ok(())
}

const fn pack_dialog_button_ids(default_id: i32, cancel_id: i32) -> usize {
    ((default_id as usize) << 16) | (cancel_id as u16 as usize)
}

fn dialog_lines(text: &str) -> Vec<String> {
    let lines = text
        .lines()
        .map(str::trim)
        .filter(|line| !line.is_empty())
        .map(ToString::to_string)
        .collect::<Vec<_>>();
    if lines.is_empty() {
        vec![text.to_string()]
    } else {
        lines
    }
}

fn dialog_list_height(line_count: usize) -> i32 {
    (line_count.max(1) as i32 * 22 + 8).clamp(40, 180)
}

fn make_access_label(text: &str) -> String {
    if text.contains('&') {
        return text.to_string();
    }
    let mut inserted = false;
    let mut result = String::with_capacity(text.len() + 1);
    for ch in text.chars() {
        if !inserted && ch.is_alphabetic() {
            result.push('&');
            inserted = true;
        }
        result.push(ch);
    }
    if inserted { result } else { format!("&{text}") }
}

unsafe fn populate_listbox_lines(hwnd: HWND, lines: &[String]) {
    SendMessageW(hwnd, LB_RESETCONTENT, 0, 0);
    for line in lines {
        let text = wide(line);
        SendMessageW(hwnd, LB_ADDSTRING, 0, text.as_ptr() as LPARAM);
    }
    if !lines.is_empty() {
        SendMessageW(hwnd, LB_SETCURSEL, 0, 0);
    }
}

fn progress_lines_text(lines: &[String], current_line: usize) -> String {
    if lines.is_empty() {
        return "Brak informacji o postępie.".to_string();
    }

    lines
        .iter()
        .enumerate()
        .map(|(index, line)| {
            if index == current_line {
                format!("> {line}")
            } else {
                format!("  {line}")
            }
        })
        .collect::<Vec<_>>()
        .join("\r\n")
}

unsafe fn set_progress_text(hwnd: HWND, lines: &[String], current_line: usize) {
    let text = progress_lines_text(lines, current_line);
    SetWindowTextW(hwnd, wide(&text).as_ptr());
}

fn progress_ratio(processed: ItemCounts, total: ItemCounts) -> f64 {
    if total.bytes > 0 {
        (processed.bytes as f64 / total.bytes as f64).clamp(0.0, 1.0)
    } else {
        let processed_items = processed.files + processed.directories;
        let total_items = total.files + total.directories;
        if total_items == 0 {
            1.0
        } else {
            (processed_items as f64 / total_items as f64).clamp(0.0, 1.0)
        }
    }
}

fn format_duration(duration: Duration) -> String {
    let total_seconds = duration.as_secs();
    let hours = total_seconds / 3600;
    let minutes = (total_seconds % 3600) / 60;
    let seconds = total_seconds % 60;
    if hours > 0 {
        format!("{hours:02}:{minutes:02}:{seconds:02}")
    } else {
        format!("{minutes:02}:{seconds:02}")
    }
}

fn format_rate(bytes_per_second: f64) -> String {
    if !bytes_per_second.is_finite() || bytes_per_second <= 0.0 {
        "0 B/s".to_string()
    } else {
        format!("{}/s", format_bytes(bytes_per_second.round() as u64))
    }
}

fn build_progress_lines(
    action: &str,
    status: &str,
    current: Option<&Path>,
    destination: Option<&Path>,
    processed: ItemCounts,
    total: ItemCounts,
    elapsed: Duration,
) -> Vec<String> {
    let percent = progress_ratio(processed, total) * 100.0;
    let elapsed_seconds = elapsed.as_secs_f64();
    let speed = if elapsed_seconds > 0.0 {
        processed.bytes as f64 / elapsed_seconds
    } else {
        0.0
    };
    let remaining = total.bytes.saturating_sub(processed.bytes);
    let eta = if remaining > 0 && speed > 0.0 {
        Some(Duration::from_secs_f64(remaining as f64 / speed))
    } else {
        None
    };
    let mut lines = vec![
        format!("Operacja: {action}"),
        status.to_string(),
        format!("Procent: {percent:.0}%"),
        format!(
            "Dane: {} z {}",
            format_bytes(processed.bytes),
            format_bytes(total.bytes)
        ),
        format!("Prędkość: {}", format_rate(speed)),
        format!("Czas trwania: {}", format_duration(elapsed)),
        format!(
            "Pozostało: {}",
            eta.map(format_duration)
                .unwrap_or_else(|| "nieznane".to_string())
        ),
        format!("Pliki: {} z {}", processed.files, total.files),
        format!(
            "Katalogi: {} z {}",
            processed.directories, total.directories
        ),
    ];
    if let Some(path) = current {
        lines.insert(2, format!("Bieżący element: {}", path.display()));
    }
    if let Some(path) = destination {
        lines.insert(3, format!("Cel: {}", path.display()));
    }
    lines
}

fn protect_secret(secret: &str) -> io::Result<String> {
    if secret.is_empty() {
        return Ok(String::new());
    }
    let mut input_bytes = secret.as_bytes().to_vec();
    let input = CRYPT_INTEGER_BLOB {
        cbData: input_bytes.len() as u32,
        pbData: input_bytes.as_mut_ptr(),
    };
    let mut output = CRYPT_INTEGER_BLOB {
        cbData: 0,
        pbData: null_mut(),
    };
    let ok = unsafe {
        CryptProtectData(
            &input,
            null(),
            null(),
            null(),
            null(),
            CRYPTPROTECT_UI_FORBIDDEN,
            &mut output,
        )
    };
    if ok == 0 {
        let error_code = unsafe { GetLastError() };
        return Err(io::Error::other(format!(
            "nie udało się zaszyfrować danych: {}",
            error_code
        )));
    }
    let bytes = unsafe { std::slice::from_raw_parts(output.pbData, output.cbData as usize) };
    let encoded = hex_encode(bytes);
    unsafe {
        LocalFree(output.pbData as _);
    }
    Ok(encoded)
}

fn unprotect_secret(secret: &str) -> Option<String> {
    if secret.is_empty() {
        return Some(String::new());
    }
    let mut encrypted = hex_decode(secret).ok()?;
    let input = CRYPT_INTEGER_BLOB {
        cbData: encrypted.len() as u32,
        pbData: encrypted.as_mut_ptr(),
    };
    let mut output = CRYPT_INTEGER_BLOB {
        cbData: 0,
        pbData: null_mut(),
    };
    let ok = unsafe {
        CryptUnprotectData(
            &input,
            null_mut(),
            null(),
            null(),
            null(),
            CRYPTPROTECT_UI_FORBIDDEN,
            &mut output,
        )
    };
    if ok == 0 {
        return None;
    }
    let bytes = unsafe { std::slice::from_raw_parts(output.pbData, output.cbData as usize) };
    let text = String::from_utf8(bytes.to_vec()).ok();
    unsafe {
        LocalFree(output.pbData as _);
    }
    text
}

fn hex_encode(bytes: &[u8]) -> String {
    const HEX: &[u8; 16] = b"0123456789abcdef";
    let mut output = String::with_capacity(bytes.len() * 2);
    for byte in bytes {
        output.push(HEX[(byte >> 4) as usize] as char);
        output.push(HEX[(byte & 0x0f) as usize] as char);
    }
    output
}

fn hex_decode(text: &str) -> Result<Vec<u8>, ()> {
    if !text.len().is_multiple_of(2) {
        return Err(());
    }
    let mut output = Vec::with_capacity(text.len() / 2);
    let bytes = text.as_bytes();
    let mut index = 0usize;
    while index < bytes.len() {
        let high = decode_hex_nibble(bytes[index])?;
        let low = decode_hex_nibble(bytes[index + 1])?;
        output.push((high << 4) | low);
        index += 2;
    }
    Ok(output)
}

fn decode_hex_nibble(value: u8) -> Result<u8, ()> {
    match value {
        b'0'..=b'9' => Ok(value - b'0'),
        b'a'..=b'f' => Ok(value - b'a' + 10),
        b'A'..=b'F' => Ok(value - b'A' + 10),
        _ => Err(()),
    }
}

fn settings_file_path() -> PathBuf {
    let base = std::env::var_os("APPDATA")
        .map(PathBuf::from)
        .or_else(|| std::env::current_dir().ok())
        .unwrap_or_else(|| PathBuf::from("."));
    base.join("AmigaFmNative").join("settings.json")
}

fn load_settings() -> AppSettings {
    let path = settings_file_path();
    match fs::read_to_string(path) {
        Ok(contents) => serde_json::from_str::<PersistedAppSettings>(&contents)
            .map(|persisted| AppSettings {
                view_options: persisted.view_options,
                favorite_directories: persisted.favorite_directories,
                favorite_files: persisted.favorite_files,
                network_resources: persisted
                    .network_resources
                    .into_iter()
                    .map(PersistedNetworkResource::into_runtime)
                    .collect(),
                discovery_cache: DiscoveryCache::from_persisted(persisted.discovery_cache),
            })
            .unwrap_or_default(),
        Err(_) => AppSettings::default(),
    }
}

fn save_settings_file(settings: &AppSettings) -> io::Result<()> {
    let path = settings_file_path();
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent)?;
    }
    let persisted = PersistedAppSettings {
        view_options: settings.view_options,
        favorite_directories: settings.favorite_directories.clone(),
        favorite_files: settings.favorite_files.clone(),
        network_resources: settings
            .network_resources
            .iter()
            .map(PersistedNetworkResource::from_runtime)
            .collect::<io::Result<Vec<_>>>()?,
        discovery_cache: settings.discovery_cache.to_persisted(),
    };
    let contents = serde_json::to_string_pretty(&persisted)
        .map_err(|error| io::Error::other(error.to_string()))?;
    fs::write(path, contents)
}

fn build_network_uri(
    scheme: &str,
    host: &str,
    directory: &str,
    resource: &NetworkResource,
) -> String {
    let auth = if resource.anonymous {
        String::new()
    } else if !resource.username.trim().is_empty() {
        format!("{}@", resource.username.trim())
    } else {
        String::new()
    };
    if directory.is_empty() {
        format!("{scheme}://{auth}{host}")
    } else {
        format!("{scheme}://{auth}{host}/{directory}")
    }
}

fn encode_uri_path(path: &str) -> String {
    let normalized = path.replace('\\', "/");
    let mut encoded = String::with_capacity(normalized.len());
    for byte in normalized.bytes() {
        match byte {
            b'A'..=b'Z' | b'a'..=b'z' | b'0'..=b'9' | b'-' | b'_' | b'.' | b'~' | b'/' | b':' => {
                encoded.push(byte as char)
            }
            _ => {
                encoded.push('%');
                encoded.push_str(&format!("{byte:02X}"));
            }
        }
    }
    encoded
}

fn is_media_extension(path: &Path) -> bool {
    const MEDIA_EXTENSIONS: &[&str] = &[
        "mp3", "wav", "flac", "ogg", "m4a", "aac", "wma", "opus", "mp4", "m4v", "mkv", "avi",
        "mov", "wmv", "webm", "mpeg", "mpg", "ts", "m2ts", "flv",
    ];
    path.extension()
        .and_then(|extension| extension.to_str())
        .map(|extension| MEDIA_EXTENSIONS.contains(&extension.to_ascii_lowercase().as_str()))
        .unwrap_or(false)
}

fn can_stream_remote_media(resource: &NetworkResource) -> bool {
    match resource.protocol {
        NetworkProtocol::WebDav | NetworkProtocol::Ftp | NetworkProtocol::Ftps => {
            resource.anonymous || resource.password.is_empty()
        }
        _ => false,
    }
}

fn remote_media_stream_target(resource: &NetworkResource, path: &Path) -> Option<String> {
    if !is_media_extension(path) || !can_stream_remote_media(resource) {
        return None;
    }

    let remote_path = encode_uri_path(&path.to_string_lossy());
    let auth = if resource.anonymous {
        String::new()
    } else if !resource.username.trim().is_empty() {
        format!("{}@", resource.username.trim())
    } else {
        String::new()
    };

    match resource.protocol {
        NetworkProtocol::Ftp => Some(format!(
            "ftp://{}{}/{}",
            auth,
            resource.normalized_host(),
            remote_path.trim_start_matches('/')
        )),
        NetworkProtocol::Ftps => Some(format!(
            "ftps://{}{}/{}",
            auth,
            resource.normalized_host(),
            remote_path.trim_start_matches('/')
        )),
        NetworkProtocol::WebDav => {
            let base_url = normalize_webdav_base_url(&resource.host);
            let (scheme, rest) = if let Some(rest) = base_url.strip_prefix("https://") {
                ("https", rest)
            } else if let Some(rest) = base_url.strip_prefix("http://") {
                ("http", rest)
            } else {
                return None;
            };
            Some(format!(
                "{scheme}://{}{rest}{}",
                auth,
                if remote_path.starts_with('/') {
                    remote_path
                } else {
                    format!("/{remote_path}")
                }
            ))
        }
        _ => None,
    }
}

fn read_drive_entries() -> Vec<PanelEntry> {
    let mut entries = vec![
        PanelEntry::favorite_directories_root(),
        PanelEntry::favorite_files_root(),
    ];
    for letter in b'A'..=b'Z' {
        let drive = format!("{}:\\", letter as char);
        let path = PathBuf::from(&drive);
        if path.exists() {
            entries.push(PanelEntry::drive(path));
        }
    }
    entries.push(PanelEntry::network_placeholder());
    entries
}

fn read_network_resource_entries(resources: &[NetworkResource]) -> Vec<PanelEntry> {
    let mut entries = resources
        .iter()
        .filter(|resource| !resource.host.trim().is_empty())
        .cloned()
        .map(PanelEntry::network_resource)
        .collect::<Vec<_>>();
    entries.sort_by(|left, right| left.name.to_lowercase().cmp(&right.name.to_lowercase()));
    entries
}

fn read_favorite_entries(
    favorites: &[PathBuf],
    directories_only: bool,
    view_options: ViewOptions,
) -> Vec<PanelEntry> {
    let mut entries = favorites
        .iter()
        .filter_map(|path| {
            let entry = PanelEntry::from_path(path.clone(), view_options).ok()?;
            match (directories_only, entry.kind) {
                (true, EntryKind::Directory) | (false, EntryKind::File) => Some(entry),
                _ => None,
            }
        })
        .collect::<Vec<_>>();
    entries.sort_by(|left, right| left.name.to_lowercase().cmp(&right.name.to_lowercase()));
    entries
}

fn open_network_resource(resource: &NetworkResource) -> io::Result<()> {
    let target = resource.launch_target();
    open_with_system_target(&target)
}

fn remote_error_to_io(error: remotefs::RemoteError) -> io::Error {
    let message = error.to_string();
    if is_permission_denied_message(&message) {
        io::Error::new(io::ErrorKind::PermissionDenied, message)
    } else {
        io::Error::other(message)
    }
}

fn is_permission_denied_message(message: &str) -> bool {
    let lowered = message.to_lowercase();
    lowered.contains("permission denied")
        || lowered.contains("access denied")
        || lowered.contains("odmowa dostępu")
}

fn is_permission_denied_error(error: &io::Error) -> bool {
    error.kind() == io::ErrorKind::PermissionDenied
        || is_permission_denied_message(&error.to_string())
}

fn is_not_found_message(message: &str) -> bool {
    let lowered = message.to_lowercase();
    lowered.contains("no such file")
        || lowered.contains("not found")
        || lowered.contains("nie znaleziono")
        || lowered.contains("nie istnieje")
}

fn is_not_found_error(error: &io::Error) -> bool {
    error.kind() == io::ErrorKind::NotFound || is_not_found_message(&error.to_string())
}

fn is_auth_failure_message(message: &str) -> bool {
    let lowered = message.to_lowercase();
    lowered.contains("authentication failed")
        || lowered.contains("auth failed")
        || lowered.contains("authentication error")
        || lowered.contains("userauth")
        || lowered.contains("invalid password")
        || lowered.contains("incorrect password")
        || lowered.contains("wrong password")
        || lowered.contains("bad authentication")
        || lowered.contains("invalid credentials")
}

fn is_sudo_auth_failure_message(message: &str) -> bool {
    let lowered = message.to_lowercase();
    lowered.contains("sudo:")
        || lowered.contains("try again")
        || lowered.contains("a password is required")
        || lowered.contains("nieprawidłowe hasło sudo")
        || lowered.contains("sudo odrzuciło uwierzytelnienie")
        || lowered.contains("polecenie sudo nie powiodło się")
}

fn is_sftp_retry_message(message: &str) -> bool {
    is_permission_denied_message(message)
        || is_auth_failure_message(message)
        || is_sudo_auth_failure_message(message)
}

fn is_sftp_retry_error(error: &io::Error) -> bool {
    is_permission_denied_error(error)
        || is_auth_failure_message(&error.to_string())
        || is_sudo_auth_failure_message(&error.to_string())
}

fn is_sftp_source_retry_message(message: &str) -> bool {
    message.to_lowercase().starts_with("źródło sftp:")
}

fn is_sftp_destination_retry_message(message: &str) -> bool {
    message.to_lowercase().starts_with("cel sftp:")
}

fn is_uac_canceled_error(error: &io::Error) -> bool {
    matches!(error.raw_os_error(), Some(1223))
}

fn normalize_remote_directory(value: &str) -> PathBuf {
    let trimmed = value.trim();
    if trimmed.is_empty() {
        PathBuf::from("/")
    } else if trimmed.starts_with('/') {
        PathBuf::from(trimmed)
    } else {
        Path::new("/").join(trimmed)
    }
}

fn shell_quote_posix(value: &str) -> String {
    format!("'{}'", value.replace('\'', "'\"'\"'"))
}

fn remote_shell_path(path: &Path) -> String {
    path.to_string_lossy().replace('\\', "/")
}

fn remote_child_path(base: &Path, name: &str) -> PathBuf {
    let base = remote_shell_path(base);
    if base == "/" || base.is_empty() {
        PathBuf::from(format!("/{name}"))
    } else {
        PathBuf::from(format!("{}/{}", base.trim_end_matches('/'), name))
    }
}

fn is_sftp_sudo_location(remote: &RemoteLocation) -> bool {
    remote.uses_sftp_sudo()
}

fn exec_remote_command(resource: &NetworkResource, cmd: &str) -> io::Result<(u32, String)> {
    let mut client = connect_remote_client(resource)?;
    let result = client.exec(cmd);
    client.disconnect();
    result
}

fn exec_sftp_sudo_command(resource: &NetworkResource, cmd: &str) -> io::Result<String> {
    if resource.protocol != NetworkProtocol::Sftp {
        return Err(io::Error::other("sudo jest obsługiwane tylko dla SFTP"));
    }

    let path_prefix = "export PATH=/usr/local/sbin:/usr/local/bin:/usr/sbin:/usr/bin:/sbin:/bin;";
    let sudo_cmd = if resource.sudo_password.is_empty() {
        format!(
            "{path_prefix} sudo -n sh -c {} 2>&1",
            shell_quote_posix(cmd)
        )
    } else {
        format!(
            "{path_prefix} printf '%s\\n' {} | sudo -S -k -p '' sh -c {} 2>&1",
            shell_quote_posix(&resource.sudo_password),
            shell_quote_posix(cmd)
        )
    };
    let (sudo_code, sudo_output) = exec_remote_command(resource, &sudo_cmd)?;
    if sudo_code == 0 {
        return Ok(sudo_output.trim().to_string());
    }

    if !resource.root_password.is_empty() {
        let su_cmd = format!(
            "{path_prefix} printf '%s\\n' {} | su root -c {} 2>&1",
            shell_quote_posix(&resource.root_password),
            shell_quote_posix(cmd)
        );
        let (su_code, su_output) = exec_remote_command(resource, &su_cmd)?;
        if su_code == 0 {
            return Ok(su_output.trim().to_string());
        }
        let combined = format!("sudo: {}; su: {}", sudo_output.trim(), su_output.trim());
        return Err(io::Error::new(
            io::ErrorKind::PermissionDenied,
            combined
                .trim_matches(|ch: char| ch.is_whitespace() || ch == ';')
                .to_string(),
        ));
    }

    Err(io::Error::new(
        io::ErrorKind::PermissionDenied,
        if sudo_output.trim().is_empty() {
            "nieprawidłowe hasło sudo albo sudo odrzuciło uwierzytelnienie".to_string()
        } else {
            sudo_output
        },
    ))
}

fn stat_sftp_path_via_sudo(remote: &RemoteLocation, path: &Path) -> io::Result<(bool, u64)> {
    let target = shell_quote_posix(&remote_shell_path(path));
    let output = exec_sftp_sudo_command(
        &remote.resource,
        &format!(
            "if [ -d {target} ]; then printf 'D\\t0\\n'; \
             elif [ -f {target} ]; then size=$(wc -c < {target} 2>/dev/null || echo 0); printf 'F\\t%s\\n' \"$size\"; \
             elif [ -e {target} ]; then printf 'F\\t0\\n'; \
             else exit 2; fi"
        ),
    )?;
    let mut parts = output.trim().split('\t');
    let kind = parts.next().unwrap_or("");
    let size = parts
        .next()
        .and_then(|value| value.trim().parse::<u64>().ok())
        .unwrap_or(0);
    match kind {
        "D" => Ok((true, 0)),
        "F" => Ok((false, size)),
        _ => Err(io::Error::other(
            "nie udało się odczytać typu elementu przez sudo",
        )),
    }
}

fn list_sftp_children_via_sudo(
    remote: &RemoteLocation,
    path: &Path,
) -> io::Result<Vec<(String, PathBuf, bool, u64)>> {
    let dir = shell_quote_posix(&remote_shell_path(path));
    let output = exec_sftp_sudo_command(
        &remote.resource,
        &format!(
            "dir={dir}; \
             for f in \"$dir\"/* \"$dir\"/.[!.]* \"$dir\"/..?*; do \
                 [ -e \"$f\" ] || continue; \
                 name=$(basename \"$f\"); \
                 if [ -d \"$f\" ]; then \
                     printf 'D\\t0\\t%s\\n' \"$name\"; \
                 else \
                     size=$(wc -c < \"$f\" 2>/dev/null || echo 0); \
                     printf 'F\\t%s\\t%s\\n' \"$size\" \"$name\"; \
                 fi; \
             done"
        ),
    )?;
    let mut entries = Vec::new();
    for line in output.lines() {
        let mut parts = line.splitn(3, '\t');
        let kind = parts.next().unwrap_or("").trim();
        let size = parts
            .next()
            .and_then(|value| value.trim().parse::<u64>().ok())
            .unwrap_or(0);
        let name = parts.next().unwrap_or("").trim().to_string();
        if name.is_empty() {
            continue;
        }
        entries.push((
            name.clone(),
            remote_child_path(path, &name),
            kind == "D",
            size,
        ));
    }
    entries.sort_by(|left, right| match (left.2, right.2) {
        (true, false) => std::cmp::Ordering::Less,
        (false, true) => std::cmp::Ordering::Greater,
        _ => left.0.to_lowercase().cmp(&right.0.to_lowercase()),
    });
    Ok(entries)
}

fn read_sftp_entries_via_sudo(
    remote: &RemoteLocation,
    view_options: ViewOptions,
) -> io::Result<Vec<PanelEntry>> {
    let mut entries = list_sftp_children_via_sudo(remote, &remote.path)?
        .into_iter()
        .map(|(name, path, is_dir, size)| PanelEntry {
            name: name.clone(),
            path: Some(path),
            link_target: None,
            network_resource: None,
            kind: if is_dir {
                EntryKind::Directory
            } else {
                EntryKind::File
            },
            size_bytes: if view_options.show_size && !is_dir {
                Some(size)
            } else {
                None
            },
            type_label: if is_dir {
                Some("katalog".to_string())
            } else {
                Some(infer_file_type(&name))
            },
            created_label: None,
            modified_label: None,
        })
        .collect::<Vec<_>>();
    entries.sort_by(|left, right| match (left.kind, right.kind) {
        (EntryKind::Directory, EntryKind::File) => std::cmp::Ordering::Less,
        (EntryKind::File, EntryKind::Directory) => std::cmp::Ordering::Greater,
        _ => left.name.to_lowercase().cmp(&right.name.to_lowercase()),
    });
    Ok(entries)
}

fn download_sftp_file_via_sudo_to_temp(
    remote: &RemoteLocation,
    path: &Path,
    progress: &mut ProgressReporter,
) -> io::Result<PathBuf> {
    let size = exec_sftp_sudo_command(
        &remote.resource,
        &format!("wc -c < {}", shell_quote_posix(&remote_shell_path(path))),
    )
    .ok()
    .and_then(|output| output.trim().parse::<u64>().ok())
    .unwrap_or(0);
    progress.total = ItemCounts {
        files: 1,
        directories: 0,
        bytes: size,
    };

    let file_name = path
        .file_name()
        .map(|item| item.to_string_lossy().to_string())
        .filter(|name| !name.is_empty())
        .unwrap_or_else(|| "remote.bin".to_string());
    let local_path = unique_temp_file_path(
        "AmigaFmNativeRemote",
        &format!("{}_{}", remote.resource.effective_display_name(), file_name),
    )?;

    let remote_temp = exec_remote_command(&remote.resource, "mktemp")
        .map(|(_, output)| output.trim().to_string())
        .and_then(|output| {
            if output.is_empty() {
                Err(io::Error::other(
                    "nie udało się przygotować pliku tymczasowego na serwerze",
                ))
            } else {
                Ok(PathBuf::from(output))
            }
        })?;

    progress.update(
        "Pobieranie pliku z użyciem sudo.",
        Some(path),
        Some(&local_path),
    );
    let source = shell_quote_posix(&remote_shell_path(path));
    let temp = shell_quote_posix(&remote_shell_path(&remote_temp));
    exec_sftp_sudo_command(
        &remote.resource,
        &format!("cp -- {source} {temp} && chmod 0644 {temp}"),
    )?;

    let mut client = connect_remote_client(&remote.resource)?;
    let result = client.download_to(&remote_temp, Box::new(fs::File::create(&local_path)?));
    let _ = client.remove_file(&remote_temp);
    client.disconnect();
    match result {
        Ok(_) => {
            progress.add_counts(
                ItemCounts {
                    files: 1,
                    directories: 0,
                    bytes: size,
                },
                "Plik pobrany.",
                Some(path),
                Some(&local_path),
            );
            Ok(local_path)
        }
        Err(error) if is_canceled_io_error(&error) || progress.is_canceled() => {
            let _ = fs::remove_file(&local_path);
            Err(io::Error::new(io::ErrorKind::Interrupted, "anulowano"))
        }
        Err(error) => {
            let _ = fs::remove_file(&local_path);
            Err(error)
        }
    }
}

fn remote_parent_path(path: &Path) -> Option<PathBuf> {
    let normalized = if path.as_os_str().is_empty() {
        Path::new("/")
    } else {
        path
    };
    let parent = normalized.parent()?;
    if parent == normalized || parent.as_os_str().is_empty() {
        None
    } else {
        Some(parent.to_path_buf())
    }
}

fn remote_display_suffix(path: &Path) -> String {
    let normalized = if path.as_os_str().is_empty() {
        Path::new("/")
    } else {
        path
    };
    if normalized == Path::new("/") {
        String::new()
    } else {
        format!(" [{}]", normalized.display())
    }
}

fn parse_host_and_port(host: &str, default_port: u16) -> (String, u16) {
    let trimmed = host
        .trim()
        .trim_start_matches("ftp://")
        .trim_start_matches("ftps://")
        .trim_start_matches("sftp://")
        .trim_start_matches("http://")
        .trim_start_matches("https://");
    if let Some((name, port)) = trimmed.rsplit_once(':')
        && let Ok(port) = port.parse::<u16>()
    {
        return (name.to_string(), port);
    }
    (trimmed.to_string(), default_port)
}

fn normalize_webdav_base_url(host: &str) -> String {
    let trimmed = host.trim();
    if trimmed.starts_with("http://") || trimmed.starts_with("https://") {
        trimmed.trim_end_matches('/').to_string()
    } else {
        format!("http://{}", trimmed.trim_end_matches('/'))
    }
}

fn nfs_auth_credential() -> opaque_auth<'static> {
    let auth = auth_unix {
        stamp: 0xaaaa_aaaa,
        machinename: Opaque::borrowed(b"unknown"),
        uid: 0xffff_fffe,
        gid: 0xffff_fffe,
        gids: vec![],
    };
    opaque_auth::auth_unix(&auth)
}

fn resolve_nfs_host(host: &str) -> io::Result<String> {
    let socket = resolve_first_socket_address(&format!("{host}:111"))
        .ok_or_else(|| io::Error::other("nie udało się rozwiązać hosta NFS"))?;
    Ok(socket.ip().to_string())
}

fn normalize_nfs_mount_path(path: &str) -> String {
    let trimmed = path.trim();
    if trimmed.is_empty() {
        "/".to_string()
    } else if trimmed.starts_with('/') {
        trimmed.to_string()
    } else {
        format!("/{trimmed}")
    }
}

fn format_nfs_exports(exports: &[String]) -> String {
    let mut preview = exports.iter().take(5).cloned().collect::<Vec<_>>();
    if exports.len() > 5 {
        preview.push("...".to_string());
    }
    preview.join(", ")
}

fn list_nfs_exports(resolved_host: &str, portmapper_port: u16) -> io::Result<Vec<String>> {
    let runtime = tokio::runtime::Builder::new_current_thread()
        .enable_all()
        .build()
        .map_err(|error| io::Error::other(error.to_string()))?;
    runtime.block_on(async move {
        let ip = resolved_host
            .parse()
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error))?;
        let mut portmapper = PortmapperClient::new(TokioIo::new(
            TokioTcpStream::connect(SocketAddr::new(ip, portmapper_port)).await?,
        ));
        let mount_port = portmapper
            .getport(
                nfs3_client::nfs3_types::mount::PROGRAM,
                nfs3_client::nfs3_types::mount::VERSION,
            )
            .await
            .map_err(nfs_error_to_io)?;
        let mut mount_client = MountClient::new_with_auth(
            TokioIo::new(TokioTcpStream::connect(SocketAddr::new(ip, mount_port)).await?),
            nfs_auth_credential(),
            opaque_auth::default(),
        );
        let exports: exports<'static, 'static> =
            mount_client.export().await.map_err(nfs_error_to_io)?;
        let mut result = Vec::new();
        let mut seen = HashSet::new();
        for export in exports.into_inner() {
            let path = normalize_nfs_mount_path(&String::from_utf8_lossy(export.ex_dir.0.as_ref()));
            if seen.insert(path.clone()) {
                result.push(path);
            }
        }
        result.sort_by_key(|export| export.to_lowercase());
        Ok(result)
    })
}

fn nfs_result<T, E: std::fmt::Debug>(
    result: nfs3::Nfs3Result<T, E>,
    context: &str,
) -> io::Result<T> {
    match result {
        nfs3::Nfs3Result::Ok(value) => Ok(value),
        nfs3::Nfs3Result::Err((status, details)) => Err(io::Error::other(format!(
            "{context}: {status} ({details:?})"
        ))),
    }
}

fn nfs_metadata_to_remote(attrs: &nfs3::fattr3) -> RemoteMetadata {
    let file_type = match attrs.type_ {
        nfs3::ftype3::NF3DIR => RemoteFileType::Directory,
        nfs3::ftype3::NF3LNK => RemoteFileType::Symlink,
        _ => RemoteFileType::File,
    };
    RemoteMetadata::default()
        .file_type(file_type)
        .size(attrs.size)
        .modified(SystemTime::from(attrs.mtime))
}

async fn nfs_lookup_path(connection: &mut NfsConnection, path: &Path) -> io::Result<nfs3::nfs_fh3> {
    let normalized = if path.as_os_str().is_empty() {
        Path::new("/")
    } else {
        path
    };
    if normalized == Path::new("/") {
        return Ok(connection.root_nfs_fh3());
    }
    let mut handle = connection.root_nfs_fh3();
    for component in normalized.components() {
        match component {
            std::path::Component::RootDir | std::path::Component::CurDir => {}
            std::path::Component::Normal(name) => {
                let name = name.to_string_lossy().into_owned();
                let lookup = connection
                    .lookup(&nfs3::LOOKUP3args {
                        what: nfs3::diropargs3 {
                            dir: handle.clone(),
                            name: filename3::from(name.as_bytes()),
                        },
                    })
                    .await
                    .map_err(nfs_error_to_io)?;
                handle = nfs_result(lookup, "lookup NFS")?.object;
            }
            _ => return Err(io::Error::other("nieobsługiwana ścieżka NFS")),
        }
    }
    Ok(handle)
}

async fn nfs_list_dir(connection: &mut NfsConnection, path: &Path) -> io::Result<Vec<RemoteFile>> {
    let dir_handle = nfs_lookup_path(connection, path).await?;
    let mut cookie = nfs3::cookie3::default();
    let mut cookieverf = nfs3::cookieverf3::default();
    let mut entries = Vec::new();
    loop {
        let readdirplus = connection
            .readdirplus(&nfs3::READDIRPLUS3args {
                dir: dir_handle.clone(),
                cookie,
                cookieverf,
                maxcount: 128 * 1024,
                dircount: 128 * 1024,
            })
            .await
            .map_err(nfs_error_to_io)?;
        let readdirplus = nfs_result(readdirplus, "odczyt katalogu NFS")?;
        let chunk = readdirplus.reply.entries.into_inner();
        cookie = chunk
            .last()
            .map(|entry| entry.cookie)
            .unwrap_or_else(nfs3::cookie3::default);
        cookieverf = readdirplus.cookieverf;
        for entry in chunk {
            let name = String::from_utf8_lossy(entry.name.as_ref()).into_owned();
            if name == "." || name == ".." {
                continue;
            }
            let metadata = match entry.name_attributes {
                Nfs3Option::Some(attrs) => nfs_metadata_to_remote(&attrs),
                Nfs3Option::None => RemoteMetadata::default(),
            };
            let metadata = if metadata.is_symlink() {
                match entry.name_handle {
                    Nfs3Option::Some(handle) => match nfs_readlink(connection, handle).await {
                        Ok(target) => metadata.symlink(target),
                        Err(_) => metadata,
                    },
                    Nfs3Option::None => metadata,
                }
            } else {
                metadata
            };
            let full_path = if path == Path::new("/") {
                Path::new("/").join(&name)
            } else {
                path.join(&name)
            };
            entries.push(RemoteFile {
                path: full_path,
                metadata,
            });
        }
        if readdirplus.reply.eof {
            break;
        }
    }
    Ok(entries)
}

async fn nfs_stat(connection: &mut NfsConnection, path: &Path) -> io::Result<RemoteFile> {
    let handle = nfs_lookup_path(connection, path).await?;
    let attrs = connection
        .getattr(&nfs3::GETATTR3args {
            object: handle.clone(),
        })
        .await
        .map_err(nfs_error_to_io)?;
    let attrs = nfs_result(attrs, "pobieranie atrybutów NFS")?.obj_attributes;
    let mut metadata = nfs_metadata_to_remote(&attrs);
    if metadata.is_symlink() {
        if let Ok(target) = nfs_readlink(connection, handle).await {
            metadata = metadata.symlink(target);
        }
    }
    Ok(RemoteFile {
        path: if path.as_os_str().is_empty() {
            PathBuf::from("/")
        } else {
            path.to_path_buf()
        },
        metadata,
    })
}

async fn nfs_readlink(
    connection: &mut NfsConnection,
    handle: nfs3::nfs_fh3,
) -> io::Result<PathBuf> {
    let result = connection
        .readlink(&nfs3::READLINK3args { symlink: handle })
        .await
        .map_err(nfs_error_to_io)?;
    let result = nfs_result(result, "odczyt linku NFS")?;
    Ok(PathBuf::from(
        String::from_utf8_lossy(result.data.as_ref()).into_owned(),
    ))
}

async fn nfs_exists(connection: &mut NfsConnection, path: &Path) -> io::Result<bool> {
    match nfs_lookup_path(connection, path).await {
        Ok(_) => Ok(true),
        Err(error) => {
            let message = error.to_string();
            if message.contains("NFS3ERR_NOENT") || message.contains("NFS3ERR_NOTDIR") {
                Ok(false)
            } else {
                Err(error)
            }
        }
    }
}

async fn nfs_create_dir(connection: &mut NfsConnection, path: &Path) -> io::Result<()> {
    let parent = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("/"));
    let name = path
        .file_name()
        .map(|name| name.to_string_lossy().into_owned())
        .ok_or_else(|| io::Error::other("nie można utworzyć katalogu NFS"))?;
    let parent_handle = nfs_lookup_path(connection, parent).await?;
    let result = connection
        .mkdir(&nfs3::MKDIR3args {
            where_: nfs3::diropargs3 {
                dir: parent_handle,
                name: filename3::from(name.as_bytes()),
            },
            attributes: nfs3::sattr3::default(),
        })
        .await
        .map_err(nfs_error_to_io)?;
    let _ = nfs_result(result, "tworzenie katalogu NFS")?;
    Ok(())
}

async fn nfs_rename(connection: &mut NfsConnection, src: &Path, dest: &Path) -> io::Result<()> {
    let src_parent = src
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("/"));
    let dest_parent = dest
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("/"));
    let src_name = src
        .file_name()
        .map(|name| name.to_string_lossy().into_owned())
        .ok_or_else(|| io::Error::other("nie można określić źródła NFS"))?;
    let dest_name = dest
        .file_name()
        .map(|name| name.to_string_lossy().into_owned())
        .ok_or_else(|| io::Error::other("nie można określić celu NFS"))?;
    let src_parent_handle = nfs_lookup_path(connection, src_parent).await?;
    let dest_parent_handle = nfs_lookup_path(connection, dest_parent).await?;
    let result = connection
        .rename(&nfs3::RENAME3args {
            from: nfs3::diropargs3 {
                dir: src_parent_handle,
                name: filename3::from(src_name.as_bytes()),
            },
            to: nfs3::diropargs3 {
                dir: dest_parent_handle,
                name: filename3::from(dest_name.as_bytes()),
            },
        })
        .await
        .map_err(nfs_error_to_io)?;
    let _ = nfs_result(result, "zmiana nazwy NFS")?;
    Ok(())
}

async fn nfs_remove_file(connection: &mut NfsConnection, path: &Path) -> io::Result<()> {
    let parent = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("/"));
    let name = path
        .file_name()
        .map(|name| name.to_string_lossy().into_owned())
        .ok_or_else(|| io::Error::other("nie można usunąć pliku NFS"))?;
    let parent_handle = nfs_lookup_path(connection, parent).await?;
    let result = connection
        .remove(&nfs3::REMOVE3args {
            object: nfs3::diropargs3 {
                dir: parent_handle,
                name: filename3::from(name.as_bytes()),
            },
        })
        .await
        .map_err(nfs_error_to_io)?;
    let _ = nfs_result(result, "usuwanie pliku NFS")?;
    Ok(())
}

async fn nfs_remove_dir(connection: &mut NfsConnection, path: &Path) -> io::Result<()> {
    let parent = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("/"));
    let name = path
        .file_name()
        .map(|name| name.to_string_lossy().into_owned())
        .ok_or_else(|| io::Error::other("nie można usunąć katalogu NFS"))?;
    let parent_handle = nfs_lookup_path(connection, parent).await?;
    let result = connection
        .rmdir(&nfs3::RMDIR3args {
            object: nfs3::diropargs3 {
                dir: parent_handle,
                name: filename3::from(name.as_bytes()),
            },
        })
        .await
        .map_err(nfs_error_to_io)?;
    let _ = nfs_result(result, "usuwanie katalogu NFS")?;
    Ok(())
}

async fn nfs_download_file(
    connection: &mut NfsConnection,
    src: &Path,
    dest: &mut dyn Write,
    mut progress: Option<&mut ProgressReporter>,
) -> io::Result<u64> {
    let handle = nfs_lookup_path(connection, src).await?;
    let mut offset = 0u64;
    let mut written = 0u64;
    loop {
        let read = connection
            .read(&nfs3::READ3args {
                file: handle.clone(),
                offset,
                count: 128 * 1024,
            })
            .await
            .map_err(nfs_error_to_io)?;
        let read = nfs_result(read, "odczyt pliku NFS")?;
        let chunk = read.data.0;
        if chunk.is_empty() {
            break;
        }
        dest.write_all(chunk.as_ref())?;
        written += chunk.len() as u64;
        offset += chunk.len() as u64;
        if let Some(progress) = progress.as_deref_mut() {
            progress.add_counts(
                ItemCounts {
                    files: 0,
                    directories: 0,
                    bytes: chunk.len() as u64,
                },
                "Trwa kopiowanie danych.",
                Some(src),
                None,
            );
        }
        if read.eof {
            break;
        }
    }
    Ok(written)
}

async fn nfs_create_file_handle(
    connection: &mut NfsConnection,
    path: &Path,
) -> io::Result<nfs3::nfs_fh3> {
    let parent = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("/"));
    let name = path
        .file_name()
        .map(|name| name.to_string_lossy().into_owned())
        .ok_or_else(|| io::Error::other("nie można utworzyć pliku NFS"))?;
    let parent_handle = nfs_lookup_path(connection, parent).await?;
    let created = connection
        .create(&nfs3::CREATE3args {
            where_: nfs3::diropargs3 {
                dir: parent_handle,
                name: filename3::from(name.as_bytes()),
            },
            how: nfs3::createhow3::UNCHECKED(nfs3::sattr3::default()),
        })
        .await
        .map_err(nfs_error_to_io)?;
    let created = nfs_result(created, "tworzenie pliku NFS")?;
    match created.obj {
        Nfs3Option::Some(handle) => Ok(handle),
        Nfs3Option::None => Err(io::Error::other("serwer NFS nie zwrócił uchwytu pliku")),
    }
}

async fn nfs_upload_file(
    connection: &mut NfsConnection,
    path: &Path,
    _size_hint: u64,
    reader: &mut dyn Read,
    mut progress: Option<&mut ProgressReporter>,
) -> io::Result<u64> {
    let file_handle = nfs_create_file_handle(connection, path).await?;
    let mut offset = 0u64;
    let mut buffer = vec![0u8; 128 * 1024];
    loop {
        let read = reader.read(&mut buffer)?;
        if read == 0 {
            break;
        }
        let write = connection
            .write(&nfs3::WRITE3args {
                file: file_handle.clone(),
                offset,
                count: read as u32,
                stable: nfs3::stable_how::FILE_SYNC,
                data: Opaque::borrowed(&buffer[..read]),
            })
            .await
            .map_err(nfs_error_to_io)?;
        let write = nfs_result(write, "zapis pliku NFS")?;
        offset += write.count as u64;
        if let Some(progress) = progress.as_deref_mut() {
            progress.add_counts(
                ItemCounts {
                    files: 0,
                    directories: 0,
                    bytes: write.count as u64,
                },
                "Trwa kopiowanie danych.",
                None,
                Some(path),
            );
        }
    }
    Ok(offset)
}

fn nfs_error_to_io(error: nfs3_client::error::Error) -> io::Error {
    io::Error::other(error.to_string())
}

fn connect_remote_client(resource: &NetworkResource) -> io::Result<RemoteClient> {
    match resource.protocol {
        NetworkProtocol::Ftp | NetworkProtocol::Ftps => {
            let default_port = if resource.protocol == NetworkProtocol::Ftps {
                990
            } else {
                21
            };
            let (host, port) = parse_host_and_port(&resource.host, default_port);
            if host.is_empty() {
                return Err(io::Error::other("brak hosta połączenia"));
            }
            let mut client = FtpFs::new(&host, port);
            if !resource.anonymous && !resource.username.trim().is_empty() {
                client = client.username(resource.username.trim());
            }
            if !resource.anonymous && !resource.password.is_empty() {
                client = client.password(&resource.password);
            }
            if resource.protocol == NetworkProtocol::Ftps {
                client = client.secure(false, false);
            }
            let mut client = RemoteClient::Ftp(client);
            client.connect()?;
            Ok(client)
        }
        NetworkProtocol::Sftp => {
            let (host, port) = parse_host_and_port(&resource.host, 22);
            if host.is_empty() {
                return Err(io::Error::other("brak hosta połączenia"));
            }
            let runtime = Arc::new(
                tokio::runtime::Builder::new_current_thread()
                    .enable_all()
                    .build()
                    .map_err(|error| io::Error::other(error.to_string()))?,
            );
            let mut opts = SshOpts::new(&host)
                .port(port)
                .connection_timeout(Duration::from_secs(12));
            if !resource.username.trim().is_empty() {
                opts = opts.username(resource.username.trim());
            }
            if !resource.password.is_empty() {
                opts = opts.password(&resource.password);
            }
            if !resource.ssh_key.trim().is_empty() {
                opts = opts.key_storage(Box::new(FixedSshKeyStorage {
                    path: PathBuf::from(resource.ssh_key.trim()),
                }));
            }
            let mut client = RemoteClient::Sftp(SftpFs::russh(opts, runtime));
            client.connect()?;
            Ok(client)
        }
        NetworkProtocol::WebDav => {
            let base_url = normalize_webdav_base_url(&resource.host);
            if base_url == "http://" {
                return Err(io::Error::other("brak hosta połączenia"));
            }
            let username = if resource.anonymous {
                ""
            } else {
                resource.username.trim()
            };
            let password = if resource.anonymous {
                ""
            } else {
                resource.password.as_str()
            };
            let mut client = RemoteClient::WebDav(WebDAVFs::new(username, password, &base_url));
            client.connect()?;
            Ok(client)
        }
        NetworkProtocol::Smb => Err(io::Error::other(
            "SMB jest obsługiwane przez ścieżkę sieciową Windows",
        )),
        NetworkProtocol::Nfs => Ok(RemoteClient::Nfs(NfsSession::new(resource.clone())?)),
    }
}

fn read_remote_entries(
    remote: &RemoteLocation,
    view_options: ViewOptions,
) -> io::Result<Vec<PanelEntry>> {
    if is_sftp_sudo_location(remote) {
        return read_sftp_entries_via_sudo(remote, view_options);
    }
    let mut client = connect_remote_client(&remote.resource)?;
    let mut entries = client
        .list_dir(&remote.path)?
        .into_iter()
        .filter(|file| {
            let name = file.name();
            name != "." && name != ".."
        })
        .map(|file| PanelEntry::from_remote_file(file, view_options))
        .collect::<Vec<_>>();
    client.disconnect();
    entries.sort_by(|left, right| match (left.kind, right.kind) {
        (EntryKind::Directory, EntryKind::File) => std::cmp::Ordering::Less,
        (EntryKind::File, EntryKind::Directory) => std::cmp::Ordering::Greater,
        _ => left.name.to_lowercase().cmp(&right.name.to_lowercase()),
    });
    Ok(entries)
}

#[derive(Clone)]
struct ArchiveListItem {
    path: PathBuf,
    is_dir: bool,
    size: Option<u64>,
    modified_label: Option<String>,
}

#[derive(Clone)]
struct ArchiveExtractJob {
    archive_path: PathBuf,
    items: Vec<PathBuf>,
    destination: PathBuf,
}

enum ProcessOutputEvent {
    Stdout(Vec<u8>),
    Stderr(Vec<u8>),
}

fn seven_zip_executable() -> PathBuf {
    for candidate in [
        PathBuf::from(r"C:\Program Files\7-Zip\7z.exe"),
        PathBuf::from(r"C:\Program Files (x86)\7-Zip\7z.exe"),
    ] {
        if candidate.exists() {
            return candidate;
        }
    }
    PathBuf::from("7z.exe")
}

fn run_7z<I, S>(args: I) -> io::Result<Output>
where
    I: IntoIterator<Item = S>,
    S: AsRef<OsStr>,
{
    Command::new(seven_zip_executable())
        .args(args)
        .creation_flags(CREATE_NO_WINDOW_FLAG)
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .output()
}

fn run_7z_checked<I, S>(args: I) -> io::Result<String>
where
    I: IntoIterator<Item = S>,
    S: AsRef<OsStr>,
{
    let output = run_7z(args)?;
    if output.status.success() {
        Ok(String::from_utf8_lossy(&output.stdout).into_owned())
    } else {
        let stderr = String::from_utf8_lossy(&output.stderr);
        let stdout = String::from_utf8_lossy(&output.stdout);
        Err(io::Error::other(if !stderr.trim().is_empty() {
            stderr.trim().to_string()
        } else if !stdout.trim().is_empty() {
            stdout.trim().to_string()
        } else {
            format!("7-Zip zakończył pracę kodem {}", output.status)
        }))
    }
}

fn spawn_process_output_reader<R>(
    mut reader: R,
    sender: Sender<ProcessOutputEvent>,
    stderr: bool,
) -> thread::JoinHandle<()>
where
    R: Read + Send + 'static,
{
    thread::spawn(move || {
        let mut buffer = [0u8; 4096];
        loop {
            let Ok(read) = reader.read(&mut buffer) else {
                break;
            };
            if read == 0 {
                break;
            }
            let chunk = buffer[..read].to_vec();
            let event = if stderr {
                ProcessOutputEvent::Stderr(chunk)
            } else {
                ProcessOutputEvent::Stdout(chunk)
            };
            if sender.send(event).is_err() {
                break;
            }
        }
    })
}

fn last_7z_progress_ratio(text: &str) -> Option<f64> {
    for (percent_index, _) in text.match_indices('%').rev() {
        let before = &text[..percent_index];
        let mut digits_reversed = String::new();
        let mut found_digit = false;
        for ch in before.chars().rev() {
            if ch.is_ascii_digit() {
                digits_reversed.push(ch);
                found_digit = true;
                continue;
            }
            if found_digit {
                break;
            }
            if ch.is_whitespace() || ch == '\r' || ch == '\n' {
                continue;
            }
            break;
        }
        if digits_reversed.is_empty() {
            continue;
        }
        let digits = digits_reversed.chars().rev().collect::<String>();
        if let Ok(value) = digits.parse::<u8>() {
            if value <= 100 {
                return Some(value as f64 / 100.0);
            }
        }
    }
    None
}

fn trim_progress_tail(text: &mut String) {
    const MAX_CHARS: usize = 2048;
    if text.chars().count() <= MAX_CHARS {
        return;
    }
    let tail = text
        .chars()
        .rev()
        .take(MAX_CHARS)
        .collect::<String>()
        .chars()
        .rev()
        .collect::<String>();
    *text = tail;
}

fn update_7z_progress_from_chunk(
    chunk: &[u8],
    progress_tail: &mut String,
    progress: &mut ProgressReporter,
    status: &str,
    current: Option<&Path>,
    destination: Option<&Path>,
    base_counts: ItemCounts,
    phase_counts: ItemCounts,
) {
    progress_tail.push_str(&String::from_utf8_lossy(chunk));
    trim_progress_tail(progress_tail);
    if let Some(ratio) = last_7z_progress_ratio(progress_tail) {
        let counts = base_counts.added(phase_counts.scaled(ratio));
        progress.set_processed(counts, status, current, destination);
    }
}

fn run_7z_checked_with_progress<I, S>(
    args: I,
    progress: &mut ProgressReporter,
    status: &str,
    current: Option<&Path>,
    destination: Option<&Path>,
    base_counts: ItemCounts,
    phase_counts: ItemCounts,
) -> io::Result<String>
where
    I: IntoIterator<Item = S>,
    S: AsRef<OsStr>,
{
    let mut child = Command::new(seven_zip_executable())
        .args(args)
        .creation_flags(CREATE_NO_WINDOW_FLAG)
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()?;

    let stdout = child
        .stdout
        .take()
        .ok_or_else(|| io::Error::other("nie udało się odczytać wyjścia 7-Zip"))?;
    let stderr = child
        .stderr
        .take()
        .ok_or_else(|| io::Error::other("nie udało się odczytać błędów 7-Zip"))?;
    let (sender, receiver) = mpsc::channel();
    let stdout_reader = spawn_process_output_reader(stdout, sender.clone(), false);
    let stderr_reader = spawn_process_output_reader(stderr, sender, true);

    let mut stdout_bytes = Vec::new();
    let mut stderr_bytes = Vec::new();
    let mut progress_tail = String::new();
    progress.set_processed(base_counts, status, current, destination);
    progress.force_update();

    let status_code = loop {
        while let Ok(event) = receiver.try_recv() {
            match event {
                ProcessOutputEvent::Stdout(chunk) => {
                    update_7z_progress_from_chunk(
                        &chunk,
                        &mut progress_tail,
                        progress,
                        status,
                        current,
                        destination,
                        base_counts,
                        phase_counts,
                    );
                    stdout_bytes.extend(chunk);
                }
                ProcessOutputEvent::Stderr(chunk) => stderr_bytes.extend(chunk),
            }
        }

        if progress.is_canceled() {
            let _ = child.kill();
            let _ = child.wait();
            let _ = stdout_reader.join();
            let _ = stderr_reader.join();
            return Err(io::Error::new(io::ErrorKind::Interrupted, "anulowano"));
        }

        if let Some(status) = child.try_wait()? {
            break status;
        }
        progress.send_update_if_due(false);
        thread::sleep(Duration::from_millis(50));
    };

    let _ = stdout_reader.join();
    let _ = stderr_reader.join();
    while let Ok(event) = receiver.try_recv() {
        match event {
            ProcessOutputEvent::Stdout(chunk) => {
                update_7z_progress_from_chunk(
                    &chunk,
                    &mut progress_tail,
                    progress,
                    status,
                    current,
                    destination,
                    base_counts,
                    phase_counts,
                );
                stdout_bytes.extend(chunk);
            }
            ProcessOutputEvent::Stderr(chunk) => stderr_bytes.extend(chunk),
        }
    }

    if status_code.success() {
        progress.set_processed(
            base_counts.added(phase_counts),
            status,
            current,
            destination,
        );
        progress.force_update();
        Ok(String::from_utf8_lossy(&stdout_bytes).into_owned())
    } else {
        let stderr = String::from_utf8_lossy(&stderr_bytes);
        let stdout = String::from_utf8_lossy(&stdout_bytes);
        Err(io::Error::other(if !stderr.trim().is_empty() {
            stderr.trim().to_string()
        } else if !stdout.trim().is_empty() {
            stdout.trim().to_string()
        } else {
            format!("7-Zip zakończył pracę kodem {status_code}")
        }))
    }
}

fn is_archive_file_path(path: &Path) -> bool {
    let Some(name) = path
        .file_name()
        .map(|name| name.to_string_lossy().to_lowercase())
    else {
        return false;
    };
    if name.ends_with(".tar.gz")
        || name.ends_with(".tar.bz2")
        || name.ends_with(".tar.xz")
        || name.ends_with(".tar.zst")
        || name.ends_with(".tar.lzma")
        || name.ends_with(".tar.lz")
        || name.ends_with(".tar.z")
        || name.ends_with(".cpio.gz")
        || name.ends_with(".part1.rar")
    {
        return true;
    }
    const EXTENSIONS: &[&str] = &[
        "7z",
        "zip",
        "zipx",
        "rar",
        "r00",
        "tar",
        "gz",
        "gzip",
        "tgz",
        "tpz",
        "bz2",
        "bzip2",
        "tbz",
        "tbz2",
        "xz",
        "txz",
        "zst",
        "tzst",
        "z",
        "taz",
        "lzma",
        "lzma86",
        "wim",
        "swm",
        "esd",
        "ppkg",
        "cab",
        "iso",
        "udf",
        "img",
        "ima",
        "vhd",
        "vhdx",
        "avhdx",
        "vdi",
        "vmdk",
        "qcow",
        "qcow2",
        "qcow2c",
        "dmg",
        "hfs",
        "hfsx",
        "apfs",
        "fat",
        "ntfs",
        "mbr",
        "gpt",
        "simg",
        "lpimg",
        "xar",
        "pkg",
        "xip",
        "ar",
        "a",
        "arj",
        "lha",
        "lzh",
        "chm",
        "msi",
        "msp",
        "nsis",
        "deb",
        "udeb",
        "rpm",
        "cpio",
        "squashfs",
        "cramfs",
        "apk",
        "jar",
        "xpi",
        "ipa",
        "appx",
        "msix",
        "appxbundle",
        "msixbundle",
        "b64",
        "pmd",
        "001",
    ];
    path.extension()
        .and_then(|extension| extension.to_str())
        .map(|extension| EXTENSIONS.contains(&extension.to_ascii_lowercase().as_str()))
        .unwrap_or(false)
}

fn normalize_archive_inner_path(path: &Path) -> PathBuf {
    let text = path.to_string_lossy().replace('\\', "/");
    let mut parts = Vec::new();
    for part in text.split('/') {
        match part {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            value => parts.push(value.to_string()),
        }
    }
    PathBuf::from(parts.join("/"))
}

fn archive_inner_to_7z_arg(path: &Path) -> String {
    normalize_archive_inner_path(path)
        .to_string_lossy()
        .replace('\\', "/")
}

fn archive_internal_display_path(path: &Path) -> String {
    if path.as_os_str().is_empty() {
        String::new()
    } else {
        format!(": /{}", archive_inner_to_7z_arg(path))
    }
}

fn archive_entry_name(path: &Path) -> String {
    let text = archive_inner_to_7z_arg(path);
    text.rsplit('/')
        .next()
        .filter(|name| !name.is_empty())
        .unwrap_or(&text)
        .to_string()
}

fn archive_output_folder_name(path: &Path) -> String {
    let name = display_name(path);
    let lowered = name.to_lowercase();
    for suffix in [
        ".tar.gz",
        ".tar.bz2",
        ".tar.xz",
        ".tar.zst",
        ".tar.lzma",
        ".tar.lz",
        ".tar.z",
        ".part1.rar",
    ] {
        if lowered.ends_with(suffix) {
            return name[..name.len() - suffix.len()].to_string();
        }
    }
    path.file_stem()
        .map(|value| value.to_string_lossy().to_string())
        .filter(|value| !value.is_empty())
        .unwrap_or_else(|| "wypakowane".to_string())
}

fn extract_archive_items_to_dir(
    archive_path: &Path,
    items: &[PathBuf],
    destination: &Path,
) -> io::Result<()> {
    fs::create_dir_all(destination)?;
    let mut args = vec![
        OsStringLike::from("x"),
        OsStringLike::from("-y"),
        OsStringLike::from(format!("-o{}", destination.display())),
        OsStringLike::from("--"),
        OsStringLike::from_path(archive_path),
    ];
    for item in items {
        args.push(OsStringLike::from(archive_inner_to_7z_arg(item)));
    }
    run_7z_checked(args.iter().map(|arg| arg.as_os_str())).map(|_| ())
}

fn extract_archive_items_to_dir_progress(
    archive_path: &Path,
    items: &[PathBuf],
    destination: &Path,
    base_counts: ItemCounts,
    phase_counts: ItemCounts,
    progress: &mut ProgressReporter,
) -> io::Result<()> {
    fs::create_dir_all(destination)?;
    let mut args = vec![
        OsStringLike::from("x"),
        OsStringLike::from("-bsp1"),
        OsStringLike::from("-bb1"),
        OsStringLike::from("-y"),
        OsStringLike::from(format!("-o{}", destination.display())),
        OsStringLike::from("--"),
        OsStringLike::from_path(archive_path),
    ];
    for item in items {
        args.push(OsStringLike::from(archive_inner_to_7z_arg(item)));
    }
    run_7z_checked_with_progress(
        args.iter().map(|arg| arg.as_os_str()),
        progress,
        "Trwa wypakowywanie archiwum.",
        Some(archive_path),
        Some(destination),
        base_counts,
        phase_counts,
    )
    .map(|_| ())
}

fn delete_archive_items(archive_path: &Path, items: &[PathBuf]) -> io::Result<()> {
    if items.is_empty() {
        return Err(io::Error::other("brak elementów do usunięcia"));
    }
    let mut args = vec![
        OsStringLike::from("d"),
        OsStringLike::from("-y"),
        OsStringLike::from("--"),
        OsStringLike::from_path(archive_path),
    ];
    for item in items {
        args.push(OsStringLike::from(archive_inner_to_7z_arg(item)));
    }
    run_7z_checked(args.iter().map(|arg| arg.as_os_str())).map(|_| ())
}

fn add_files_to_archive(
    archive_path: &Path,
    targets: &[PathBuf],
    base_dir: &Path,
) -> io::Result<()> {
    if targets.is_empty() {
        return Err(io::Error::other("brak elementów do dodania"));
    }
    let mut command = Command::new(seven_zip_executable());
    command.current_dir(base_dir);
    command.creation_flags(CREATE_NO_WINDOW_FLAG);
    command.arg("a").arg("--").arg(archive_path);
    for target in targets {
        let relative = target.strip_prefix(base_dir).unwrap_or(target.as_path());
        command.arg(relative);
    }
    let output = command
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .output()?;
    if output.status.success() {
        Ok(())
    } else {
        let stderr = String::from_utf8_lossy(&output.stderr);
        let stdout = String::from_utf8_lossy(&output.stdout);
        Err(io::Error::other(if !stderr.trim().is_empty() {
            stderr.trim().to_string()
        } else {
            stdout.trim().to_string()
        }))
    }
}

fn default_archive_name(targets: &[PathBuf]) -> String {
    match targets {
        [single] => format!("{}.7z", display_name(single)),
        [] => "archiwum.7z".to_string(),
        _ => "archiwum.7z".to_string(),
    }
}

fn ensure_archive_create_extension(mut path: PathBuf, format: ArchiveFormat) -> PathBuf {
    let Some(file_name) = path
        .file_name()
        .map(|value| value.to_string_lossy().to_string())
    else {
        return path;
    };
    let desired_suffix = format!(".{}", format.extension());
    let lowered = file_name.to_ascii_lowercase();
    if lowered.ends_with(&desired_suffix) {
        return path;
    }
    for known in ArchiveFormat::ALL {
        let known_suffix = format!(".{}", known.extension());
        if lowered.ends_with(&known_suffix) {
            let prefix_len = file_name.len().saturating_sub(known_suffix.len());
            path.set_file_name(format!("{}{}", &file_name[..prefix_len], desired_suffix));
            return path;
        }
    }
    path.set_file_name(format!("{file_name}{desired_suffix}"));
    path
}

fn archive_output_path(base_dir: &Path, options: &ArchiveCreateOptions) -> PathBuf {
    let mut path = PathBuf::from(options.name.trim());
    if path.as_os_str().is_empty() {
        path = PathBuf::from(format!("archiwum.{}", options.format.extension()));
    }
    path = ensure_archive_create_extension(path, options.format);
    if path.is_absolute() {
        path
    } else {
        base_dir.join(path)
    }
}

fn create_archive_with_7z_progress(
    output_path: &Path,
    targets: &[PathBuf],
    options: &ArchiveCreateOptions,
    source_counts: ItemCounts,
    progress: &mut ProgressReporter,
) -> io::Result<()> {
    if targets.is_empty() {
        return Err(io::Error::other("brak elementów do archiwizacji"));
    }
    if options.encrypted && !options.format.supports_encryption() {
        return Err(io::Error::other(
            "wybrany format nie obsługuje szyfrowania przez 7-Zip",
        ));
    }
    if let Some(parent) = output_path.parent() {
        fs::create_dir_all(parent)?;
    }
    if options.format.compressed_tar_codec().is_some() {
        return create_compressed_tar_archive_progress(
            output_path,
            targets,
            options,
            source_counts,
            progress,
        );
    }
    if options.format.is_single_file_compressor() {
        let [single] = targets else {
            return Err(io::Error::other(
                "wybrany format kompresuje pojedynczy plik. Do wielu elementów użyj tar.gz, tar.bz2, tar.xz, 7z albo zip",
            ));
        };
        if single.is_dir() {
            return Err(io::Error::other(
                "wybrany format kompresuje pojedynczy plik. Do katalogu użyj tar.gz, tar.bz2, tar.xz, 7z albo zip",
            ));
        }
    }
    create_direct_7z_archive_progress(
        output_path,
        targets,
        options,
        ItemCounts::default(),
        source_counts,
        progress,
        "Trwa tworzenie archiwum.",
    )
}

fn archive_create_progress_total(source_counts: ItemCounts, format: ArchiveFormat) -> ItemCounts {
    if format.compressed_tar_codec().is_some() {
        ItemCounts {
            bytes: source_counts.bytes.saturating_mul(2),
            ..source_counts
        }
    } else {
        source_counts
    }
}

fn create_direct_7z_archive_progress(
    output_path: &Path,
    targets: &[PathBuf],
    options: &ArchiveCreateOptions,
    base_counts: ItemCounts,
    phase_counts: ItemCounts,
    progress: &mut ProgressReporter,
    status: &str,
) -> io::Result<()> {
    let mut args = vec![
        OsStringLike::from("a"),
        OsStringLike::from("-bsp1"),
        OsStringLike::from("-bb1"),
        OsStringLike::from(format!("-t{}", options.format.seven_zip_type())),
        OsStringLike::from(format!("-mx={}", options.compression_level)),
    ];
    if options.encrypted {
        args.push(OsStringLike::from(format!("-p{}", options.password)));
        if options.format == ArchiveFormat::SevenZip {
            args.push(OsStringLike::from("-mhe=on"));
        }
        if options.format == ArchiveFormat::Zip && options.encryption.contains("AES") {
            args.push(OsStringLike::from("-mem=AES256"));
        }
    }
    if !options.volume_size.trim().is_empty() {
        args.push(OsStringLike::from(format!(
            "-v{}",
            options.volume_size.trim()
        )));
    }
    args.push(OsStringLike::from("--"));
    args.push(OsStringLike::from_path(output_path));
    for target in targets {
        args.push(OsStringLike::from_path(target));
    }
    run_7z_checked_with_progress(
        args.iter().map(|arg| arg.as_os_str()),
        progress,
        status,
        targets.first().map(PathBuf::as_path),
        Some(output_path),
        base_counts,
        phase_counts,
    )
    .map(|_| ())
}

fn create_compressed_tar_archive_progress(
    output_path: &Path,
    targets: &[PathBuf],
    options: &ArchiveCreateOptions,
    source_counts: ItemCounts,
    progress: &mut ProgressReporter,
) -> io::Result<()> {
    let codec = options
        .format
        .compressed_tar_codec()
        .ok_or_else(|| io::Error::other("nieobsługiwany format archiwum tar"))?;
    let temp_root = unique_temp_file_path("AmigaFmNativeArchiveCreate", "tar")?;
    fs::create_dir_all(&temp_root)?;
    let tar_name = compressed_tar_inner_name(output_path);
    let temp_tar = temp_root.join(tar_name);
    let result = (|| {
        let tar_options = ArchiveCreateOptions {
            format: ArchiveFormat::Tar,
            name: String::new(),
            compression_level: 0,
            encrypted: false,
            encryption: String::new(),
            password: String::new(),
            volume_size: String::new(),
        };
        create_direct_7z_archive_progress(
            &temp_tar,
            targets,
            &tar_options,
            ItemCounts::default(),
            source_counts,
            progress,
            "Trwa tworzenie tymczasowego archiwum tar.",
        )?;

        let mut compress_args = vec![
            OsStringLike::from("a"),
            OsStringLike::from("-bsp1"),
            OsStringLike::from("-bb1"),
            OsStringLike::from(format!("-t{codec}")),
            OsStringLike::from(format!("-mx={}", options.compression_level)),
        ];
        if !options.volume_size.trim().is_empty() {
            compress_args.push(OsStringLike::from(format!(
                "-v{}",
                options.volume_size.trim()
            )));
        }
        compress_args.push(OsStringLike::from("--"));
        compress_args.push(OsStringLike::from_path(output_path));
        compress_args.push(OsStringLike::from_path(&temp_tar));
        run_7z_checked_with_progress(
            compress_args.iter().map(|arg| arg.as_os_str()),
            progress,
            "Trwa kompresowanie archiwum tar.",
            Some(&temp_tar),
            Some(output_path),
            source_counts,
            ItemCounts {
                files: 0,
                directories: 0,
                bytes: source_counts.bytes,
            },
        )
        .map(|_| ())
    })();
    let _ = fs::remove_dir_all(&temp_root);
    result
}

fn compressed_tar_inner_name(output_path: &Path) -> String {
    let name = display_name(output_path);
    let lowered = name.to_ascii_lowercase();
    for suffix in [".gz", ".bz2", ".xz"] {
        if lowered.ends_with(suffix) {
            let prefix_len = name.len().saturating_sub(suffix.len());
            return name[..prefix_len].to_string();
        }
    }
    format!("{name}.tar")
}

fn joined_archive_default_name(first_part: &Path) -> String {
    let name = display_name(first_part);
    let lowered = name.to_lowercase();
    for suffix in [".001", ".002", ".003"] {
        if lowered.ends_with(suffix) {
            return name[..name.len() - suffix.len()].to_string();
        }
    }
    format!("{name}.joined")
}

fn join_split_files(parts: &[PathBuf], output: &Path) -> io::Result<()> {
    if parts.len() < 2 {
        return Err(io::Error::other("zaznacz co najmniej dwie części"));
    }
    if let Some(parent) = output.parent() {
        fs::create_dir_all(parent)?;
    }
    let mut writer = fs::File::create(output)?;
    let mut buffer = vec![0u8; 1024 * 1024];
    for part in parts {
        let mut reader = fs::File::open(part)?;
        loop {
            let read = reader.read(&mut buffer)?;
            if read == 0 {
                break;
            }
            writer.write_all(&buffer[..read])?;
        }
    }
    Ok(())
}

#[derive(Clone)]
struct ChecksumVerifyReport;

impl ChecksumVerifyReport {
    fn success_message(&self) -> String {
        "Suma kontrolna jest poprawna".to_string()
    }
}

fn checksum_output_path(base_dir: &Path, targets: &[PathBuf]) -> PathBuf {
    match targets {
        [single] => {
            let parent = single.parent().unwrap_or(base_dir);
            parent.join(format!("{}.sha256", display_name(single)))
        }
        _ => base_dir.join("checksums.sha256"),
    }
}

fn is_sha256_checksum_path(path: &Path) -> bool {
    path.extension()
        .and_then(|extension| extension.to_str())
        .map(|extension| extension.eq_ignore_ascii_case("sha256"))
        .unwrap_or(false)
}

fn resolve_checksum_file_for_target(target: &Path) -> Result<PathBuf, String> {
    if is_sha256_checksum_path(target) {
        return Ok(target.to_path_buf());
    }
    let parent = target.parent().unwrap_or_else(|| Path::new("."));
    let candidate = parent.join(format!("{}.sha256", display_name(target)));
    if candidate.exists() {
        Ok(candidate)
    } else {
        Err(format!(
            "Nie znaleziono pliku sumy kontrolnej. Oczekiwany plik to {}. Możesz też zaznaczyć bezpośrednio plik .sha256.",
            candidate.display()
        ))
    }
}

fn collect_checksum_files(path: &Path, out: &mut Vec<PathBuf>) -> io::Result<()> {
    let metadata = fs::metadata(path)?;
    if metadata.is_file() {
        out.push(path.to_path_buf());
    } else if metadata.is_dir() {
        for child in fs::read_dir(path)? {
            collect_checksum_files(&child?.path(), out)?;
        }
    }
    Ok(())
}

fn compute_sha256_with_7z(path: &Path) -> io::Result<String> {
    let output = run_7z_checked([
        OsStr::new("h"),
        OsStr::new("-scrcSHA256"),
        OsStr::new("--"),
        path.as_os_str(),
    ])?;
    for line in output.lines() {
        let token = line.split_whitespace().next().unwrap_or("");
        if token.len() == 64 && token.chars().all(|ch| ch.is_ascii_hexdigit()) {
            return Ok(token.to_ascii_lowercase());
        }
    }
    Err(io::Error::other("nie udało się odczytać sumy SHA-256"))
}

fn create_sha256_file(targets: &[PathBuf], output: &Path) -> io::Result<()> {
    let mut files = Vec::new();
    for target in targets {
        collect_checksum_files(target, &mut files)?;
    }
    if files.is_empty() {
        return Err(io::Error::other("brak plików do policzenia sumy"));
    }
    let base = output.parent().unwrap_or_else(|| Path::new("."));
    let mut contents = String::new();
    for file in files {
        let hash = compute_sha256_with_7z(&file)?;
        let name = file
            .strip_prefix(base)
            .unwrap_or(file.as_path())
            .to_string_lossy()
            .replace('\\', "/");
        contents.push_str(&format!("{hash} *{name}\n"));
    }
    fs::write(output, contents)
}

fn verify_sha256_file(checksum_file: &Path) -> io::Result<ChecksumVerifyReport> {
    let contents = fs::read_to_string(checksum_file).map_err(|error| {
        io::Error::other(format!(
            "Nie można odczytać pliku sumy kontrolnej {}: {error}",
            checksum_file.display()
        ))
    })?;
    let base = checksum_file.parent().unwrap_or_else(|| Path::new("."));
    let mut checked = 0usize;
    let mut failed = Vec::new();
    let mut missing = Vec::new();
    let mut unreadable = Vec::new();
    for line in contents.lines() {
        let trimmed = line.trim();
        if trimmed.is_empty() || trimmed.starts_with('#') {
            continue;
        }
        let mut parts = trimmed.splitn(2, char::is_whitespace);
        let expected = parts.next().unwrap_or("").to_ascii_lowercase();
        let name = parts.next().unwrap_or("").trim().trim_start_matches('*');
        if expected.len() != 64 || name.is_empty() {
            continue;
        }
        let path = base.join(name);
        if !path.exists() {
            missing.push(name.to_string());
            continue;
        }
        let actual = match compute_sha256_with_7z(&path) {
            Ok(hash) => hash,
            Err(error) => {
                unreadable.push(format!("{name}: {error}"));
                continue;
            }
        };
        checked += 1;
        if actual != expected {
            failed.push(name.to_string());
        }
    }
    if checked == 0 && missing.is_empty() && unreadable.is_empty() {
        return Err(io::Error::other(
            "Ten plik nie zawiera sum SHA-256. Wybierz plik .sha256 albo plik, obok którego znajduje się odpowiadający mu plik .sha256.",
        ));
    }
    if !failed.is_empty() || !missing.is_empty() || !unreadable.is_empty() {
        let mut parts = Vec::new();
        if !failed.is_empty() {
            parts.push("Suma kontrolna jest niezgodna".to_string());
        }
        if !missing.is_empty() {
            parts.push(format!("Nie znaleziono plików: {}", missing.join(", ")));
        }
        if !unreadable.is_empty() {
            parts.push(format!(
                "Nie udało się odczytać plików: {}",
                unreadable.join(", ")
            ));
        }
        return Err(io::Error::other(parts.join(". ")));
    }
    Ok(ChecksumVerifyReport)
}

struct OsStringLike(std::ffi::OsString);

impl OsStringLike {
    fn from<T: Into<std::ffi::OsString>>(value: T) -> Self {
        Self(value.into())
    }

    fn from_path(path: &Path) -> Self {
        Self(path.as_os_str().to_os_string())
    }

    fn as_os_str(&self) -> &OsStr {
        self.0.as_os_str()
    }
}

fn parse_archive_list(archive: &Path) -> io::Result<Vec<ArchiveListItem>> {
    let output = run_7z_checked([
        OsStr::new("l"),
        OsStr::new("-slt"),
        OsStr::new("--"),
        archive.as_os_str(),
    ])?;
    let mut items = Vec::new();
    let mut record = HashMap::<String, String>::new();
    let flush = |record: &mut HashMap<String, String>, items: &mut Vec<ArchiveListItem>| {
        let Some(raw_path) = record.remove("Path") else {
            record.clear();
            return;
        };
        if raw_path == archive.display().to_string() || raw_path.trim().is_empty() {
            record.clear();
            return;
        }
        let path = normalize_archive_inner_path(Path::new(&raw_path));
        if path.as_os_str().is_empty() {
            record.clear();
            return;
        }
        let is_dir = record
            .get("Folder")
            .map(|value| value.trim() == "+")
            .unwrap_or(false);
        let size = record
            .get("Size")
            .and_then(|value| value.trim().parse::<u64>().ok());
        let modified_label = record
            .get("Modified")
            .map(|value| value.trim().to_string())
            .filter(|value| !value.is_empty());
        items.push(ArchiveListItem {
            path,
            is_dir,
            size,
            modified_label,
        });
        record.clear();
    };

    for line in output.lines() {
        let line = line.trim_end();
        if line.is_empty() {
            flush(&mut record, &mut items);
            continue;
        }
        if let Some((key, value)) = line.split_once(" = ") {
            record.insert(key.to_string(), value.to_string());
        }
    }
    flush(&mut record, &mut items);
    Ok(items)
}

fn read_archive_entries(
    archive: &ArchiveLocation,
    view_options: ViewOptions,
) -> io::Result<Vec<PanelEntry>> {
    let inside = normalize_archive_inner_path(&archive.inside_path);
    let prefix = archive_inner_to_7z_arg(&inside);
    let prefix = prefix.trim_matches('/').to_string();
    let mut direct = HashMap::<String, ArchiveListItem>::new();

    for item in parse_archive_list(&archive.archive_path)? {
        let item_text = archive_inner_to_7z_arg(&item.path);
        let item_text = item_text.trim_matches('/').to_string();
        if !prefix.is_empty() {
            if item_text == prefix {
                continue;
            }
            let wanted_prefix = format!("{prefix}/");
            if !item_text.starts_with(&wanted_prefix) {
                continue;
            }
        }
        let rest = if prefix.is_empty() {
            item_text.as_str()
        } else {
            item_text[prefix.len() + 1..].trim_start_matches('/')
        };
        if rest.is_empty() {
            continue;
        }
        if let Some((first, _)) = rest.split_once('/') {
            let child_path = if prefix.is_empty() {
                PathBuf::from(first)
            } else {
                PathBuf::from(format!("{prefix}/{first}"))
            };
            direct
                .entry(archive_inner_to_7z_arg(&child_path))
                .or_insert(ArchiveListItem {
                    path: child_path,
                    is_dir: true,
                    size: None,
                    modified_label: None,
                });
        } else {
            direct.insert(
                item_text.clone(),
                ArchiveListItem {
                    path: PathBuf::from(item_text),
                    ..item
                },
            );
        }
    }

    let mut directories = Vec::new();
    let mut files = Vec::new();
    for item in direct.into_values() {
        let name = archive_entry_name(&item.path);
        let entry = PanelEntry {
            name: name.clone(),
            path: Some(item.path),
            link_target: None,
            network_resource: None,
            kind: if item.is_dir {
                EntryKind::Directory
            } else {
                EntryKind::File
            },
            size_bytes: if view_options.show_size && !item.is_dir {
                item.size
            } else {
                None
            },
            type_label: if item.is_dir {
                Some("katalog".to_string())
            } else {
                Some(infer_file_type(&name))
            },
            created_label: None,
            modified_label: if view_options.show_modified {
                item.modified_label
            } else {
                None
            },
        };
        if item.is_dir {
            directories.push(entry);
        } else {
            files.push(entry);
        }
    }
    directories.sort_by_key(|entry| entry.name.to_lowercase());
    files.sort_by_key(|entry| entry.name.to_lowercase());
    directories.extend(files);
    Ok(directories)
}

fn archive_item_matches_selection(item_path: &Path, selected_items: &[PathBuf]) -> bool {
    if selected_items.is_empty() {
        return true;
    }
    let item_text = archive_inner_to_7z_arg(item_path)
        .trim_matches('/')
        .to_string();
    selected_items.iter().any(|selected| {
        let selected_text = archive_inner_to_7z_arg(selected)
            .trim_matches('/')
            .to_string();
        item_text == selected_text || item_text.starts_with(&format!("{selected_text}/"))
    })
}

fn archive_extract_counts(
    archive_path: &Path,
    selected_items: &[PathBuf],
) -> io::Result<ItemCounts> {
    let mut counts = ItemCounts::default();
    for item in parse_archive_list(archive_path)? {
        if !archive_item_matches_selection(&item.path, selected_items) {
            continue;
        }
        if item.is_dir {
            counts.directories += 1;
        } else {
            counts.files += 1;
            counts.bytes = counts.bytes.saturating_add(item.size.unwrap_or(0));
        }
    }
    Ok(counts)
}

fn download_remote_file_to_temp_with_progress(
    remote: &RemoteLocation,
    path: &Path,
    progress: &mut ProgressReporter,
) -> io::Result<PathBuf> {
    if remote.uses_sftp_sudo() {
        return download_sftp_file_via_sudo_to_temp(remote, path, progress);
    }
    if progress.is_canceled() {
        return Err(io::Error::new(io::ErrorKind::Interrupted, "anulowano"));
    }
    let mut client = connect_remote_client(&remote.resource)?;
    let entry = client.stat(path)?;
    progress.total = ItemCounts {
        files: 1,
        directories: 0,
        bytes: entry.metadata().size,
    };
    let file_name = path
        .file_name()
        .map(|item| item.to_string_lossy().to_string())
        .filter(|name| !name.is_empty())
        .unwrap_or_else(|| "remote.bin".to_string());
    let local_path = unique_temp_file_path(
        "AmigaFmNativeRemote",
        &format!("{}_{}", remote.resource.effective_display_name(), file_name),
    )?;
    progress.update("Pobieranie pliku z zasobu.", Some(path), Some(&local_path));
    let result = client.download_to(path, Box::new(fs::File::create(&local_path)?));
    client.disconnect();
    match result {
        Ok(_) => {
            progress.add_counts(
                ItemCounts {
                    files: 1,
                    directories: 0,
                    bytes: entry.metadata().size,
                },
                "Plik pobrany.",
                Some(path),
                Some(&local_path),
            );
            Ok(local_path)
        }
        Err(error) if is_canceled_io_error(&error) || progress.is_canceled() => {
            let _ = fs::remove_file(&local_path);
            Err(io::Error::new(io::ErrorKind::Interrupted, "anulowano"))
        }
        Err(error) => {
            let _ = fs::remove_file(&local_path);
            Err(error)
        }
    }
}

fn remote_create_dir(remote: &RemoteLocation, path: &Path) -> io::Result<()> {
    if remote.uses_sftp_sudo() {
        exec_sftp_sudo_command(
            &remote.resource,
            &format!(
                "mkdir -p -- {}",
                shell_quote_posix(&remote_shell_path(path))
            ),
        )?;
        return Ok(());
    }
    let mut client = connect_remote_client(&remote.resource)?;
    let result = client.create_dir(path);
    client.disconnect();
    result
}

fn remote_rename(remote: &RemoteLocation, src: &Path, dest: &Path) -> io::Result<()> {
    if remote.uses_sftp_sudo() {
        exec_sftp_sudo_command(
            &remote.resource,
            &format!(
                "mv -- {} {}",
                shell_quote_posix(&remote_shell_path(src)),
                shell_quote_posix(&remote_shell_path(dest))
            ),
        )?;
        return Ok(());
    }
    let mut client = connect_remote_client(&remote.resource)?;
    let result = client.rename(src, dest);
    client.disconnect();
    result
}

fn remote_delete_targets(
    remote: &RemoteLocation,
    targets: &[PathBuf],
    entries: &[PanelEntry],
) -> io::Result<()> {
    if remote.uses_sftp_sudo() {
        for target in targets {
            exec_sftp_sudo_command(
                &remote.resource,
                &format!(
                    "rm -rf -- {}",
                    shell_quote_posix(&remote_shell_path(target))
                ),
            )?;
        }
        return Ok(());
    }
    let mut client = connect_remote_client(&remote.resource)?;
    for target in targets {
        let kind = entries
            .iter()
            .find(|entry| entry.path.as_deref() == Some(target.as_path()))
            .map(|entry| entry.kind);
        match kind {
            Some(EntryKind::Directory) => {
                delete_remote_path_recursive_with_client(&mut client, target)?
            }
            Some(EntryKind::File) => client.remove_file(target)?,
            _ => {}
        }
    }
    client.disconnect();
    Ok(())
}

fn remote_delete_targets_with_sftp_fallback(
    remote: &RemoteLocation,
    fallback: Option<&RemoteLocation>,
    targets: &[PathBuf],
    entries: &[PanelEntry],
) -> io::Result<()> {
    if remote.uses_sftp_sudo() {
        return remote_delete_targets(remote, targets, entries);
    }

    let mut client = connect_remote_client(&remote.resource)?;
    for target in targets {
        let kind = entries
            .iter()
            .find(|entry| entry.path.as_deref() == Some(target.as_path()))
            .map(|entry| entry.kind);
        let result = match kind {
            Some(EntryKind::Directory) => {
                delete_remote_path_recursive_with_client(&mut client, target)
            }
            Some(EntryKind::File) => client.remove_file(target),
            _ => Ok(()),
        };
        match result {
            Ok(()) => {}
            Err(error) if is_sftp_retry_error(&error) => {
                if let Some(fallback) = fallback {
                    remote_delete_targets(fallback, std::slice::from_ref(target), entries)?;
                } else {
                    client.disconnect();
                    return Err(error);
                }
            }
            Err(error) => {
                client.disconnect();
                return Err(error);
            }
        }
    }
    client.disconnect();
    Ok(())
}

fn build_discovery_progress_lines(
    status: &str,
    current_host: Option<&str>,
    processed_hosts: usize,
    total_hosts: usize,
    found_services: usize,
    cache_hits: usize,
    elapsed: Duration,
) -> Vec<String> {
    let safe_total = if total_hosts == 0 {
        processed_hosts.max(1)
    } else {
        total_hosts
    };
    let percent = ((processed_hosts as f64 / safe_total as f64) * 100.0).clamp(0.0, 100.0);
    let mut lines = vec![
        "Operacja: wyszukiwanie serwerów".to_string(),
        status.to_string(),
        format!("Postęp: {percent:.0}%"),
        format!("Hosty: {} z {}", processed_hosts, safe_total),
        format!("Trafienia z cache: {}", cache_hits),
        format!("Znalezione usługi: {}", found_services),
        format!("Czas trwania: {}", format_duration(elapsed)),
    ];
    if let Some(host) = current_host {
        lines.insert(2, format!("Aktualny host: {host}"));
    }
    lines
}

fn discover_local_servers_with_progress<F>(
    cancel_flag: Arc<AtomicBool>,
    cache: DiscoveryCache,
    on_update: &mut F,
) -> io::Result<DiscoveryScanResult>
where
    F: FnMut(Vec<String>),
{
    let started_at = Instant::now();
    let mut discovered = Vec::new();
    let mut seen = HashSet::new();
    let mut candidates = Vec::new();
    let mut candidate_seen = HashSet::new();
    let mut probed_hosts = Vec::new();
    let mut cache_hits = 0usize;
    let mut processed_hosts = 0usize;

    for host in local_discovery_hosts() {
        push_host_candidate(&mut candidates, &mut candidate_seen, &host);
    }

    on_update(build_discovery_progress_lines(
        "Pobieranie listy hostów z Windows API.",
        None,
        0,
        0,
        0,
        0,
        started_at.elapsed(),
    ));
    for server in discover_windows_network_servers()? {
        push_host_candidate(&mut candidates, &mut candidate_seen, &server.host);
        push_discovered_server(&mut discovered, &mut seen, server);
    }

    for server in parse_net_view_servers()? {
        push_host_candidate(&mut candidates, &mut candidate_seen, &server.host);
        push_discovered_server(&mut discovered, &mut seen, server);
    }
    for host in parse_arp_candidates()? {
        push_host_candidate(&mut candidates, &mut candidate_seen, &host);
    }

    let direct_total_hosts = candidates.len();
    on_update(build_discovery_progress_lines(
        "Przygotowano listę aktywnych hostów do szybkiego skanowania.",
        None,
        processed_hosts,
        direct_total_hosts,
        discovered.len(),
        cache_hits,
        started_at.elapsed(),
    ));

    let mut hosts_to_probe = Vec::new();
    for host in candidates {
        if let Some(cached) = cache.lookup_fresh(&host) {
            processed_hosts += 1;
            cache_hits += 1;
            for service in cached.services.iter().cloned() {
                push_discovered_server(&mut discovered, &mut seen, service);
            }
            on_update(build_discovery_progress_lines(
                "Wykorzystano wynik z cache.",
                Some(&host),
                processed_hosts,
                direct_total_hosts,
                discovered.len(),
                cache_hits,
                started_at.elapsed(),
            ));
        } else {
            hosts_to_probe.push(host);
        }
    }

    for chunk in hosts_to_probe.chunks(DISCOVERY_CONCURRENCY) {
        if cancel_flag.load(Ordering::Relaxed) {
            return Err(io::Error::new(io::ErrorKind::Interrupted, "anulowano"));
        }
        let (sender, receiver) = mpsc::channel();
        for host in chunk {
            let sender = sender.clone();
            let host = host.clone();
            thread::spawn(move || {
                let services = probe_host_services(&host);
                let _ = sender.send((host, services));
            });
        }
        drop(sender);

        for _ in 0..chunk.len() {
            if cancel_flag.load(Ordering::Relaxed) {
                return Err(io::Error::new(io::ErrorKind::Interrupted, "anulowano"));
            }
            if let Ok((host, services)) = receiver.recv() {
                processed_hosts += 1;
                for service in services.iter().cloned() {
                    push_discovered_server(&mut discovered, &mut seen, service);
                }
                probed_hosts.push((host.clone(), services));
                on_update(build_discovery_progress_lines(
                    "Skanowanie aktywnych hostów.",
                    Some(&host),
                    processed_hosts,
                    direct_total_hosts,
                    discovered.len(),
                    cache_hits,
                    started_at.elapsed(),
                ));
            }
        }
    }

    let should_run_fallback = discovered.len() <= DISCOVERY_MIN_RESULTS_BEFORE_FALLBACK
        || (direct_total_hosts == 0 && cache_hits == 0);
    let total_hosts = if should_run_fallback {
        direct_total_hosts + DISCOVERY_MAX_HOSTS_PER_INTERFACE
    } else {
        direct_total_hosts.max(processed_hosts)
    };

    if should_run_fallback {
        on_update(build_discovery_progress_lines(
            "Szybki skan dał mało wyników, uruchamianie skanu podsieci.",
            None,
            processed_hosts,
            total_hosts,
            discovered.len(),
            cache_hits,
            started_at.elapsed(),
        ));

        for host in parse_nbtstat_candidates()? {
            push_host_candidate(&mut hosts_to_probe, &mut candidate_seen, &host);
        }
        for host in discover_subnet_candidates(DISCOVERY_MAX_HOSTS_PER_INTERFACE) {
            push_host_candidate(&mut hosts_to_probe, &mut candidate_seen, &host);
        }

        let fallback_hosts = hosts_to_probe
            .into_iter()
            .filter(|host| {
                !probed_hosts
                    .iter()
                    .any(|(probed_host, _)| probed_host.eq_ignore_ascii_case(host))
                    && cache.lookup_fresh(host).is_none()
            })
            .collect::<Vec<_>>();

        for chunk in fallback_hosts.chunks(DISCOVERY_CONCURRENCY) {
            if cancel_flag.load(Ordering::Relaxed) {
                return Err(io::Error::new(io::ErrorKind::Interrupted, "anulowano"));
            }
            let (sender, receiver) = mpsc::channel();
            for host in chunk {
                let sender = sender.clone();
                let host = host.clone();
                thread::spawn(move || {
                    let services = probe_host_services(&host);
                    let _ = sender.send((host, services));
                });
            }
            drop(sender);

            for _ in 0..chunk.len() {
                if cancel_flag.load(Ordering::Relaxed) {
                    return Err(io::Error::new(io::ErrorKind::Interrupted, "anulowano"));
                }
                if let Ok((host, services)) = receiver.recv() {
                    processed_hosts += 1;
                    for service in services.iter().cloned() {
                        push_discovered_server(&mut discovered, &mut seen, service);
                    }
                    probed_hosts.push((host.clone(), services));
                    on_update(build_discovery_progress_lines(
                        "Dogłębne skanowanie podsieci.",
                        Some(&host),
                        processed_hosts,
                        total_hosts,
                        discovered.len(),
                        cache_hits,
                        started_at.elapsed(),
                    ));
                }
            }
        }
    }

    on_update(build_discovery_progress_lines(
        "Przeglądanie usług mDNS w sieci lokalnej.",
        None,
        processed_hosts,
        total_hosts,
        discovered.len(),
        cache_hits,
        started_at.elapsed(),
    ));
    for service in discover_mdns_services(cancel_flag.clone(), on_update)? {
        push_discovered_server(&mut discovered, &mut seen, service);
    }

    discovered.sort_by(|left, right| {
        left.host
            .to_lowercase()
            .cmp(&right.host.to_lowercase())
            .then_with(|| left.protocol.label().cmp(right.protocol.label()))
            .then_with(|| {
                left.default_directory
                    .to_lowercase()
                    .cmp(&right.default_directory.to_lowercase())
            })
    });
    Ok(DiscoveryScanResult {
        servers: discovered,
        probed_hosts,
        cache_hits,
        processed_hosts,
        total_hosts,
        elapsed: started_at.elapsed(),
    })
}

fn push_discovered_server(
    discovered: &mut Vec<DiscoveredServer>,
    seen: &mut HashSet<String>,
    server: DiscoveredServer,
) {
    let key = format!(
        "{}:{}:{}",
        server.protocol.label(),
        server.host.to_lowercase(),
        server.default_directory.to_lowercase()
    );
    if seen.insert(key.clone()) {
        discovered.push(server);
    } else if let Some(existing) = discovered.iter_mut().find(|existing| {
        format!(
            "{}:{}:{}",
            existing.protocol.label(),
            existing.host.to_lowercase(),
            existing.default_directory.to_lowercase()
        ) == key
    }) {
        if existing.resolved_name.is_none() {
            existing.resolved_name = server.resolved_name;
        }
        if existing.detail.is_none() {
            existing.detail = server.detail;
        } else if let (Some(existing_detail), Some(new_detail)) =
            (&mut existing.detail, server.detail)
        {
            if !existing_detail.contains(&new_detail) {
                existing_detail.push_str(", ");
                existing_detail.push_str(&new_detail);
            }
        }
    }
}

const MDNS_BROWSE_SPECS: &[MdnsBrowseSpec] = &[
    MdnsBrowseSpec {
        service_type: "_sftp-ssh._tcp.local.",
        protocol: NetworkProtocol::Sftp,
        detail: "wykryto przez mDNS",
    },
    MdnsBrowseSpec {
        service_type: "_ftp._tcp.local.",
        protocol: NetworkProtocol::Ftp,
        detail: "wykryto przez mDNS",
    },
    MdnsBrowseSpec {
        service_type: "_ftps._tcp.local.",
        protocol: NetworkProtocol::Ftps,
        detail: "wykryto przez mDNS",
    },
    MdnsBrowseSpec {
        service_type: "_ftp-ssl._tcp.local.",
        protocol: NetworkProtocol::Ftps,
        detail: "wykryto przez mDNS",
    },
    MdnsBrowseSpec {
        service_type: "_nfs._tcp.local.",
        protocol: NetworkProtocol::Nfs,
        detail: "wykryto przez mDNS",
    },
    MdnsBrowseSpec {
        service_type: "_smb._tcp.local.",
        protocol: NetworkProtocol::Smb,
        detail: "wykryto przez mDNS",
    },
    MdnsBrowseSpec {
        service_type: "_webdav._tcp.local.",
        protocol: NetworkProtocol::WebDav,
        detail: "wykryto przez mDNS",
    },
    MdnsBrowseSpec {
        service_type: "_webdavs._tcp.local.",
        protocol: NetworkProtocol::WebDav,
        detail: "wykryto przez mDNS",
    },
];

fn discover_mdns_services<F>(
    cancel_flag: Arc<AtomicBool>,
    on_update: &mut F,
) -> io::Result<Vec<DiscoveredServer>>
where
    F: FnMut(Vec<String>),
{
    let daemon = ServiceDaemon::new().map_err(|error| io::Error::other(error.to_string()))?;
    let mut browsers = Vec::new();
    for spec in MDNS_BROWSE_SPECS {
        if let Ok(receiver) = daemon.browse(spec.service_type) {
            browsers.push((*spec, receiver));
        }
    }
    if browsers.is_empty() {
        let _ = daemon.shutdown();
        return Ok(Vec::new());
    }

    let started_at = Instant::now();
    let mut discovered = Vec::new();
    let mut seen = HashSet::new();
    let mut resolved_services = HashSet::new();

    while started_at.elapsed() < MDNS_DISCOVERY_TIMEOUT {
        if cancel_flag.load(Ordering::Relaxed) {
            let _ = daemon.shutdown();
            return Err(io::Error::new(io::ErrorKind::Interrupted, "anulowano"));
        }

        let mut current_service = None::<String>;
        for (spec, receiver) in &browsers {
            while let Ok(event) = receiver.try_recv() {
                if let ServiceEvent::ServiceResolved(resolved) = event {
                    let instance_key = resolved.get_fullname().to_string();
                    if !resolved_services.insert(instance_key.clone()) {
                        continue;
                    }
                    current_service = Some(instance_key);
                    let mut added_any = false;
                    for address in resolved.get_addresses_v4() {
                        added_any = true;
                        push_discovered_server(
                            &mut discovered,
                            &mut seen,
                            DiscoveredServer {
                                host: address.to_string(),
                                resolved_name: normalized_mdns_hostname(resolved.get_hostname()),
                                protocol: spec.protocol,
                                default_directory: mdns_default_directory(spec.protocol, &resolved),
                                detail: Some(format!(
                                    "{}, port {}",
                                    spec.detail,
                                    resolved.get_port()
                                )),
                            },
                        );
                    }
                    if !added_any {
                        if let Some(host) = normalized_mdns_hostname(resolved.get_hostname()) {
                            push_discovered_server(
                                &mut discovered,
                                &mut seen,
                                DiscoveredServer {
                                    host,
                                    resolved_name: normalized_mdns_hostname(
                                        resolved.get_hostname(),
                                    ),
                                    protocol: spec.protocol,
                                    default_directory: mdns_default_directory(
                                        spec.protocol,
                                        &resolved,
                                    ),
                                    detail: Some(format!(
                                        "{}, port {}",
                                        spec.detail,
                                        resolved.get_port()
                                    )),
                                },
                            );
                        }
                    }
                }
            }
        }

        if !discovered.is_empty() || current_service.is_some() {
            on_update(build_discovery_progress_lines(
                "Przeglądanie usług mDNS w sieci lokalnej.",
                current_service.as_deref(),
                discovered.len(),
                discovered.len(),
                discovered.len(),
                0,
                started_at.elapsed(),
            ));
        }

        thread::sleep(Duration::from_millis(150));
    }

    let _ = daemon.shutdown();
    Ok(discovered)
}

fn mdns_default_directory(
    protocol: NetworkProtocol,
    resolved: &mdns_sd::ResolvedService,
) -> String {
    let path = resolved
        .get_property_val_str("path")
        .unwrap_or_default()
        .trim()
        .to_string();
    if path.is_empty() {
        String::new()
    } else if protocol == NetworkProtocol::Smb {
        path.trim_matches('/').replace('\\', "/")
    } else {
        path
    }
}

fn normalized_mdns_hostname(hostname: &str) -> Option<String> {
    let normalized = hostname.trim().trim_end_matches('.').to_string();
    if normalized.is_empty() {
        None
    } else {
        Some(normalized)
    }
}

fn push_host_candidate(candidates: &mut Vec<String>, seen: &mut HashSet<String>, host: &str) {
    let host = host.trim();
    if host.is_empty() {
        return;
    }
    let key = host.to_ascii_lowercase();
    if seen.insert(key) {
        candidates.push(host.to_string());
    }
}

fn discover_windows_network_servers() -> io::Result<Vec<DiscoveredServer>> {
    let mut servers = Vec::new();
    let mut buffer: *mut u8 = null_mut();
    let mut entries_read = 0u32;
    let mut total_entries = 0u32;
    let mut resume_handle = 0u32;
    loop {
        let status = unsafe {
            NetServerEnum(
                null(),
                101,
                &mut buffer,
                MAX_PREFERRED_LENGTH,
                &mut entries_read,
                &mut total_entries,
                SV_TYPE_WORKSTATION | SV_TYPE_SERVER | SV_TYPE_SERVER_NT | SV_TYPE_SERVER_UNIX,
                null(),
                &mut resume_handle,
            )
        };
        if !buffer.is_null() && entries_read > 0 {
            let entries = unsafe {
                std::slice::from_raw_parts(buffer as *const SERVER_INFO_101, entries_read as usize)
            };
            for entry in entries {
                let host = widestr_ptr_to_string(entry.sv101_name);
                if host.is_empty() {
                    continue;
                }
                let detail = if (entry.sv101_type & SV_TYPE_SERVER_UNIX) != 0 {
                    "wykryto przez NetServerEnum, serwer UNIX"
                } else if (entry.sv101_type & SV_TYPE_SERVER_NT) != 0 {
                    "wykryto przez NetServerEnum, serwer Windows"
                } else if (entry.sv101_type & SV_TYPE_SERVER) != 0 {
                    "wykryto przez NetServerEnum, serwer"
                } else {
                    "wykryto przez NetServerEnum, stacja robocza"
                };
                servers.push(DiscoveredServer {
                    host,
                    resolved_name: None,
                    protocol: NetworkProtocol::Smb,
                    default_directory: String::new(),
                    detail: Some(detail.to_string()),
                });
            }
            unsafe {
                NetApiBufferFree(buffer as _);
            }
            buffer = null_mut();
        }

        if status == 0 {
            break;
        }
        if status != 234 {
            if !buffer.is_null() {
                unsafe {
                    NetApiBufferFree(buffer as _);
                }
            }
            break;
        }
    }
    Ok(servers)
}

fn local_discovery_hosts() -> Vec<String> {
    let mut hosts = Vec::new();
    let mut seen = HashSet::new();
    if let Ok(name) = std::env::var("COMPUTERNAME") {
        push_host_candidate(&mut hosts, &mut seen, &name);
    }
    for interface in get_interfaces() {
        if !is_discovery_interface_candidate(&interface) {
            continue;
        }
        for network in interface.ipv4 {
            let address = network.addr();
            if is_private_ipv4_addr(address) {
                push_host_candidate(&mut hosts, &mut seen, &address.to_string());
            }
        }
    }
    hosts
}

fn enumerate_remote_smb_shares(host: &str) -> io::Result<Vec<String>> {
    let target = format!("\\\\{host}");
    let output =
        run_command_with_timeout("cmd", &["/C", "net view", &target], SMB_SHARE_ENUM_TIMEOUT)?;
    Ok(output
        .as_ref()
        .map(|output| parse_share_listing_lines(&String::from_utf8_lossy(&output.stdout)))
        .unwrap_or_default())
}

fn parse_share_listing_lines(text: &str) -> Vec<String> {
    let mut shares = Vec::new();
    let mut seen = HashSet::new();
    let mut in_table = false;
    for line in text.lines() {
        let trimmed = line.trim();
        if trimmed.is_empty() {
            continue;
        }
        if trimmed.chars().all(|ch| ch == '-' || ch == '=') {
            in_table = true;
            continue;
        }
        if !in_table {
            continue;
        }
        let lowered = trimmed.to_ascii_lowercase();
        if lowered.starts_with("the command completed")
            || lowered.starts_with("polecenie zosta")
            || lowered.starts_with("wystąpił błąd systemu")
        {
            break;
        }
        let Some(name) = trimmed.split_whitespace().next() else {
            continue;
        };
        if name.ends_with('$')
            || matches!(
                name.to_ascii_uppercase().as_str(),
                "IPC$" | "ADMIN$" | "PRINT$"
            )
        {
            continue;
        }
        if seen.insert(name.to_ascii_lowercase()) {
            shares.push(name.to_string());
        }
    }
    shares
}

fn run_command_with_timeout(
    program: &str,
    args: &[&str],
    timeout: Duration,
) -> io::Result<Option<Output>> {
    let mut child = Command::new(program)
        .args(args)
        .creation_flags(CREATE_NO_WINDOW_FLAG)
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()?;
    let started_at = Instant::now();
    loop {
        if child.try_wait()?.is_some() {
            return child.wait_with_output().map(Some);
        }
        if started_at.elapsed() >= timeout {
            let _ = child.kill();
            let _ = child.wait();
            return Ok(None);
        }
        thread::sleep(Duration::from_millis(50));
    }
}

fn parse_net_view_servers() -> io::Result<Vec<DiscoveredServer>> {
    let output = Command::new("cmd")
        .args(["/C", "net view"])
        .creation_flags(CREATE_NO_WINDOW_FLAG)
        .output()?;
    let text = String::from_utf8_lossy(&output.stdout);
    let mut servers = Vec::new();
    for line in text.lines() {
        let trimmed = line.trim();
        if let Some(server) = trimmed.strip_prefix("\\\\") {
            let host = server
                .split_whitespace()
                .next()
                .unwrap_or_default()
                .trim()
                .to_string();
            if !host.is_empty() {
                servers.push(DiscoveredServer {
                    host,
                    resolved_name: None,
                    protocol: NetworkProtocol::Smb,
                    default_directory: String::new(),
                    detail: Some("wykryto przez net view".to_string()),
                });
            }
        }
    }
    Ok(servers)
}

fn parse_arp_candidates() -> io::Result<Vec<String>> {
    let output = Command::new("arp")
        .arg("-a")
        .creation_flags(CREATE_NO_WINDOW_FLAG)
        .output()?;
    let text = String::from_utf8_lossy(&output.stdout);
    let ipv4 = Regex::new(r"\b(?:\d{1,3}\.){3}\d{1,3}\b")
        .map_err(|error| io::Error::other(error.to_string()))?;
    let mut hosts = Vec::new();
    let mut seen = HashSet::new();
    for capture in ipv4.find_iter(&text) {
        let candidate = capture.as_str().to_string();
        if is_private_ipv4(&candidate) && seen.insert(candidate.clone()) {
            hosts.push(candidate);
        }
    }
    hosts.sort_by_key(|host| ipv4_sort_key(host));
    Ok(hosts)
}

fn parse_nbtstat_candidates() -> io::Result<Vec<String>> {
    let mut hosts = Vec::new();
    let mut seen = HashSet::new();
    for host in parse_arp_candidates()? {
        let output = match Command::new("nbtstat")
            .args(["-A", &host])
            .creation_flags(CREATE_NO_WINDOW_FLAG)
            .output()
        {
            Ok(output) => output,
            Err(_) => continue,
        };
        let text = String::from_utf8_lossy(&output.stdout);
        if text.contains("<20>") || text.contains("UNIQUE") || text.contains("GROUP") {
            if seen.insert(host.clone()) {
                hosts.push(host);
            }
        }
    }
    Ok(hosts)
}

fn discover_subnet_candidates(max_hosts: usize) -> Vec<String> {
    let mut networks = Vec::new();
    for interface in get_interfaces() {
        if !is_discovery_interface_candidate(&interface) {
            continue;
        }
        let priority = discovery_interface_priority(&interface);
        for network in interface.ipv4 {
            let address = network.addr();
            if !is_private_ipv4_addr(address) {
                continue;
            }
            let scan_prefix = network.prefix_len().clamp(24, 30);
            let mask = ipv4_prefix_mask(scan_prefix);
            let address_u32 = u32::from(address);
            let network_u32 = address_u32 & mask;
            let broadcast_u32 = network_u32 | !mask;
            if broadcast_u32 <= network_u32 + 1 {
                continue;
            }
            let mut subnet_hosts = Vec::new();
            for candidate_u32 in (network_u32 + 1)..broadcast_u32 {
                if candidate_u32 == address_u32 {
                    continue;
                }
                let candidate = Ipv4Addr::from(candidate_u32);
                subnet_hosts.push(candidate.to_string());
                if subnet_hosts.len() >= max_hosts {
                    break;
                }
            }
            networks.push((priority, address, subnet_hosts));
        }
    }
    networks.sort_by(|left, right| left.0.cmp(&right.0).then_with(|| left.1.cmp(&right.1)));

    let mut hosts = Vec::new();
    let mut seen = HashSet::new();
    for (_, _, subnet_hosts) in networks {
        for host in subnet_hosts {
            if seen.insert(host.clone()) {
                hosts.push(host);
            }
        }
    }
    hosts
}

fn is_discovery_interface_candidate(interface: &netdev::Interface) -> bool {
    if interface.ipv4.is_empty() {
        return false;
    }
    !matches!(
        interface.oper_state,
        OperState::Down | OperState::LowerLayerDown | OperState::NotPresent
    )
}

fn discovery_interface_priority(interface: &netdev::Interface) -> u8 {
    let mut label = interface.name.to_ascii_lowercase();
    if let Some(friendly) = &interface.friendly_name {
        label.push(' ');
        label.push_str(&friendly.to_ascii_lowercase());
    }
    if let Some(description) = &interface.description {
        label.push(' ');
        label.push_str(&description.to_ascii_lowercase());
    }

    if contains_any(
        &label,
        &[
            "virtual",
            "vmware",
            "hyper-v",
            "loopback",
            "tunnel",
            "teredo",
            "isatap",
            "npcap",
            "bluetooth",
            "docker",
            "wsl",
            "vethernet",
        ],
    ) {
        3
    } else if contains_any(&label, &["ethernet", "wi-fi", "wifi", "wlan", "lan"]) {
        0
    } else {
        1
    }
}

fn contains_any(text: &str, needles: &[&str]) -> bool {
    needles.iter().any(|needle| text.contains(needle))
}

fn ipv4_sort_key(host: &str) -> u32 {
    host.parse::<Ipv4Addr>().map(u32::from).unwrap_or(u32::MAX)
}

fn ipv4_prefix_mask(prefix_len: u8) -> u32 {
    if prefix_len == 0 {
        0
    } else {
        u32::MAX << (32 - prefix_len)
    }
}

fn is_private_ipv4(candidate: &str) -> bool {
    candidate
        .parse::<Ipv4Addr>()
        .map(is_private_ipv4_addr)
        .unwrap_or(false)
}

fn is_private_ipv4_addr(candidate: Ipv4Addr) -> bool {
    let [a, b, ..] = candidate.octets();
    match a {
        10 => true,
        172 => (16..=31).contains(&b),
        192 => b == 168,
        _ => false,
    }
}

fn probe_host_services(host: &str) -> Vec<DiscoveredServer> {
    let mut services = Vec::new();
    let open_ports = probe_open_ports(host, &[22, 21, 990, 445, 139, 80, 8080, 443, 2049, 111]);

    if open_ports.contains(&22) || probe_sftp(host) {
        services.push(DiscoveredServer {
            host: host.to_string(),
            resolved_name: None,
            protocol: NetworkProtocol::Sftp,
            default_directory: String::new(),
            detail: Some("port 22".to_string()),
        });
    }
    if open_ports.contains(&21) || probe_ftp(host, 21) {
        services.push(DiscoveredServer {
            host: host.to_string(),
            resolved_name: None,
            protocol: NetworkProtocol::Ftp,
            default_directory: String::new(),
            detail: Some("port 21".to_string()),
        });
    }
    if open_ports.contains(&990) || probe_ftps(host) {
        services.push(DiscoveredServer {
            host: host.to_string(),
            resolved_name: None,
            protocol: NetworkProtocol::Ftps,
            default_directory: String::new(),
            detail: Some("port 990".to_string()),
        });
    }
    if open_ports.contains(&2049) || open_ports.contains(&111) {
        let nfs_services = probe_nfs_services(host);
        if nfs_services.is_empty() {
            services.push(DiscoveredServer {
                host: host.to_string(),
                resolved_name: None,
                protocol: NetworkProtocol::Nfs,
                default_directory: String::new(),
                detail: Some("port 2049 lub 111".to_string()),
            });
        } else {
            services.extend(nfs_services);
        }
    }
    if open_ports.contains(&445) || open_ports.contains(&139) {
        let mut added_share = false;
        if let Ok(shares) = enumerate_remote_smb_shares(host) {
            for share in shares {
                added_share = true;
                services.push(DiscoveredServer {
                    host: host.to_string(),
                    resolved_name: None,
                    protocol: NetworkProtocol::Smb,
                    default_directory: share.clone(),
                    detail: Some("udział SMB".to_string()),
                });
            }
        }
        services.push(DiscoveredServer {
            host: host.to_string(),
            resolved_name: None,
            protocol: NetworkProtocol::Smb,
            default_directory: String::new(),
            detail: if added_share {
                Some("host SMB z udziałami".to_string())
            } else if open_ports.contains(&445) && open_ports.contains(&139) {
                Some("porty 445 i 139".to_string())
            } else if open_ports.contains(&445) {
                Some("port 445".to_string())
            } else {
                Some("port 139".to_string())
            },
        });
    }
    if probe_webdav_from_open_ports(host, &open_ports) {
        services.push(DiscoveredServer {
            host: host.to_string(),
            resolved_name: None,
            protocol: NetworkProtocol::WebDav,
            default_directory: String::new(),
            detail: None,
        });
    }
    services
}

fn can_connect(host: &str, port: u16) -> bool {
    let address = format!("{host}:{port}");
    let socket = resolve_first_socket_address(&address);
    let Some(socket) = socket else {
        return false;
    };
    port_open_with_retry(&socket, 2)
}

fn port_open_with_retry(socket: &SocketAddr, attempts: usize) -> bool {
    for _ in 0..attempts.max(1) {
        if TcpStream::connect_timeout(socket, DISCOVERY_CONNECT_TIMEOUT).is_ok() {
            return true;
        }
    }
    false
}

fn probe_open_ports(host: &str, ports: &[u16]) -> HashSet<u16> {
    let mut open_ports = HashSet::new();
    for port in ports {
        let address = format!("{host}:{port}");
        let Some(socket) = resolve_first_socket_address(&address) else {
            continue;
        };
        if port_open_with_retry(&socket, 2) {
            open_ports.insert(*port);
        }
    }
    open_ports
}

fn resolve_first_socket_address(target: &str) -> Option<SocketAddr> {
    target.to_socket_addrs().ok()?.next()
}

fn probe_sftp(host: &str) -> bool {
    probe_banner(host, 22, None)
        .map(|banner| banner.starts_with("SSH-"))
        .unwrap_or(false)
}

fn probe_ftp(host: &str, port: u16) -> bool {
    probe_banner(host, port, None)
        .map(|banner| banner.starts_with("220"))
        .unwrap_or(false)
}

fn probe_ftps(host: &str) -> bool {
    probe_ftp(host, 990) || can_connect(host, 990)
}

fn probe_nfs_services(host: &str) -> Vec<DiscoveredServer> {
    let (raw_host, portmapper_port) = parse_host_and_port(host, 111);
    if !can_connect(host, 2049) && !can_connect(host, portmapper_port) {
        return Vec::new();
    }
    let resolved_host = resolve_nfs_host(&raw_host).ok();
    let exports = resolved_host
        .as_deref()
        .and_then(|resolved_host| list_nfs_exports(resolved_host, portmapper_port).ok())
        .unwrap_or_default();
    if !exports.is_empty() {
        return exports
            .into_iter()
            .map(|export| DiscoveredServer {
                host: host.to_string(),
                resolved_name: None,
                protocol: NetworkProtocol::Nfs,
                default_directory: export.clone(),
                detail: None,
            })
            .collect();
    }
    if can_connect(host, 2049) || can_connect(host, portmapper_port) {
        return vec![DiscoveredServer {
            host: host.to_string(),
            resolved_name: None,
            protocol: NetworkProtocol::Nfs,
            default_directory: String::new(),
            detail: Some("nie udało się odczytać eksportów".to_string()),
        }];
    }
    Vec::new()
}

fn probe_webdav_from_open_ports(host: &str, open_ports: &HashSet<u16>) -> bool {
    for port in [80u16, 8080, 443] {
        if open_ports.contains(&port) && probe_webdav_on_port(host, port) {
            return true;
        }
    }
    false
}

fn probe_webdav_on_port(host: &str, port: u16) -> bool {
    probe_banner(
        host,
        port,
        Some(format!(
            "OPTIONS / HTTP/1.1\r\nHost: {host}\r\nConnection: close\r\n\r\n"
        )),
    )
    .map(|text| {
        let lowered = text.to_ascii_lowercase();
        lowered.contains("\rdav:")
            || lowered.contains("\ndav:")
            || lowered.contains("ms-author-via: dav")
            || lowered.contains("allow:") && lowered.contains("propfind")
    })
    .unwrap_or(false)
}

fn probe_banner(host: &str, port: u16, request: Option<String>) -> Option<String> {
    let address = format!("{host}:{port}");
    let socket = resolve_first_socket_address(&address)?;
    let mut stream = TcpStream::connect_timeout(&socket, DISCOVERY_CONNECT_TIMEOUT).ok()?;
    let _ = stream.set_read_timeout(Some(DISCOVERY_IO_TIMEOUT));
    let _ = stream.set_write_timeout(Some(DISCOVERY_IO_TIMEOUT));
    if let Some(request) = request {
        stream.write_all(request.as_bytes()).ok()?;
    }
    let mut buffer = [0u8; 1024];
    let read = stream.read(&mut buffer).ok()?;
    if read == 0 {
        return None;
    }
    Some(String::from_utf8_lossy(&buffer[..read]).to_string())
}

fn search_entries(
    root: &Path,
    recursive: bool,
    regex: &Regex,
    view_options: ViewOptions,
) -> io::Result<Vec<PanelEntry>> {
    let mut directories = Vec::new();
    let mut files = Vec::new();
    collect_search_entries(
        root,
        recursive,
        regex,
        view_options,
        &mut directories,
        &mut files,
    )?;
    directories.sort_by(|left, right| left.name.to_lowercase().cmp(&right.name.to_lowercase()));
    files.sort_by(|left, right| left.name.to_lowercase().cmp(&right.name.to_lowercase()));
    directories.extend(files);
    Ok(directories)
}

fn collect_search_entries(
    root: &Path,
    recursive: bool,
    regex: &Regex,
    view_options: ViewOptions,
    directories: &mut Vec<PanelEntry>,
    files: &mut Vec<PanelEntry>,
) -> io::Result<()> {
    for entry in fs::read_dir(root)? {
        let entry = entry?;
        let path = entry.path();
        let panel_entry = match PanelEntry::from_path(path.clone(), view_options) {
            Ok(panel_entry) => panel_entry,
            Err(_) => continue,
        };
        if regex.is_match(&panel_entry.name) {
            if panel_entry.kind == EntryKind::Directory {
                directories.push(panel_entry.clone());
            } else if panel_entry.kind == EntryKind::File {
                files.push(panel_entry.clone());
            }
        }
        if recursive && panel_entry.kind == EntryKind::Directory {
            let _ = collect_search_entries(&path, true, regex, view_options, directories, files);
        }
    }
    Ok(())
}

fn summarize_path(path: &Path) -> io::Result<ItemCounts> {
    let metadata = fs::symlink_metadata(path)?;
    if metadata.file_type().is_symlink() || metadata.is_file() {
        return Ok(ItemCounts {
            files: 1,
            directories: 0,
            bytes: metadata.len(),
        });
    }
    if metadata.is_dir() {
        let mut counts = ItemCounts {
            files: 0,
            directories: 1,
            bytes: 0,
        };
        for child in fs::read_dir(path)? {
            counts.add(summarize_path(&child?.path())?);
        }
        return Ok(counts);
    }
    Ok(ItemCounts::default())
}

fn counts_description(counts: ItemCounts) -> String {
    match (counts.files, counts.directories) {
        (0, 0) => "brak elementów".to_string(),
        (files, 0) => pluralized_files(files),
        (0, directories) => pluralized_directories(directories),
        (files, directories) => format!(
            "{}, {}",
            pluralized_files(files),
            pluralized_directories(directories)
        ),
    }
}

fn format_bytes(bytes: u64) -> String {
    const UNITS: [&str; 5] = ["B", "KB", "MB", "GB", "TB"];
    let mut value = bytes as f64;
    let mut unit_index = 0usize;
    while value >= 1024.0 && unit_index + 1 < UNITS.len() {
        value /= 1024.0;
        unit_index += 1;
    }
    if unit_index == 0 {
        format!("{bytes} {}", UNITS[unit_index])
    } else {
        format!("{value:.1} {}", UNITS[unit_index])
    }
}

fn pluralized_elements(count: usize) -> String {
    if count == 1 {
        "1 element".to_string()
    } else if (2..=4).contains(&(count % 10)) && !(12..=14).contains(&(count % 100)) {
        format!("{count} elementy")
    } else {
        format!("{count} elementów")
    }
}

fn pluralized_files(count: usize) -> String {
    if count == 1 {
        "1 plik".to_string()
    } else if (2..=4).contains(&(count % 10)) && !(12..=14).contains(&(count % 100)) {
        format!("{count} pliki")
    } else {
        format!("{count} plików")
    }
}

fn pluralized_directories(count: usize) -> String {
    if count == 1 {
        "1 katalog".to_string()
    } else if (2..=4).contains(&(count % 10)) && !(12..=14).contains(&(count % 100)) {
        format!("{count} katalogi")
    } else {
        format!("{count} katalogów")
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum ConflictChoice {
    Replace,
    Skip,
    Cancel,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum OperationResult {
    Done,
    Skipped,
    Canceled,
}

unsafe fn copy_item(
    source: &Path,
    destination: &Path,
    summary: ItemCounts,
    progress: &mut ProgressReporter,
) -> io::Result<OperationResult> {
    if progress.is_canceled() {
        return Ok(OperationResult::Canceled);
    }
    if destination.exists() {
        match progress.ask_conflict(destination)? {
            ConflictChoice::Replace => remove_existing(destination)?,
            ConflictChoice::Skip => return Ok(OperationResult::Skipped),
            ConflictChoice::Cancel => return Ok(OperationResult::Canceled),
        }
    }

    let metadata = fs::symlink_metadata(source)?;
    if metadata.file_type().is_symlink() || metadata.is_file() {
        progress.update("Kopiowanie pliku.", Some(source), Some(destination));
        return copy_file_with_progress(source, destination, progress);
    }

    if metadata.is_dir() {
        progress.update("Tworzenie katalogu.", Some(source), Some(destination));
        fs::create_dir_all(destination)?;
        progress.add_counts(
            ItemCounts {
                files: 0,
                directories: 1,
                bytes: 0,
            },
            "Katalog utworzony.",
            Some(source),
            Some(destination),
        );
        for child in fs::read_dir(source)? {
            if progress.is_canceled() {
                return Ok(OperationResult::Canceled);
            }
            let child = child?;
            let child_source = child.path();
            let child_destination = destination.join(child.file_name());
            let child_summary = summarize_path(&child_source)?;
            match copy_item(&child_source, &child_destination, child_summary, progress)? {
                OperationResult::Done | OperationResult::Skipped => {}
                OperationResult::Canceled => return Ok(OperationResult::Canceled),
            }
        }
        return Ok(OperationResult::Done);
    }

    progress.add_counts(
        summary,
        "Element zakończony.",
        Some(source),
        Some(destination),
    );
    Ok(OperationResult::Done)
}

unsafe fn move_item(
    source: &Path,
    destination: &Path,
    summary: ItemCounts,
    progress: &mut ProgressReporter,
) -> io::Result<OperationResult> {
    if progress.is_canceled() {
        return Ok(OperationResult::Canceled);
    }
    if destination.exists() {
        match progress.ask_conflict(destination)? {
            ConflictChoice::Replace => remove_existing(destination)?,
            ConflictChoice::Skip => return Ok(OperationResult::Skipped),
            ConflictChoice::Cancel => return Ok(OperationResult::Canceled),
        }
    }

    progress.update("Przenoszenie elementu.", Some(source), Some(destination));
    match fs::rename(source, destination) {
        Ok(()) => {
            progress.add_counts(
                summary,
                "Element przeniesiony.",
                Some(source),
                Some(destination),
            );
            Ok(OperationResult::Done)
        }
        Err(_) => match copy_item(source, destination, summary, progress)? {
            OperationResult::Done => {
                delete_path_after_move(source)?;
                Ok(OperationResult::Done)
            }
            OperationResult::Skipped => Ok(OperationResult::Skipped),
            OperationResult::Canceled => Ok(OperationResult::Canceled),
        },
    }
}

unsafe fn copy_file_with_progress(
    source: &Path,
    destination: &Path,
    progress: &mut ProgressReporter,
) -> io::Result<OperationResult> {
    let mut input = fs::File::open(source)?;
    let mut output = fs::File::create(destination)?;
    let mut buffer = vec![0u8; 4 * 1024 * 1024];

    loop {
        if progress.is_canceled() {
            let _ = fs::remove_file(destination);
            return Ok(OperationResult::Canceled);
        }
        let read = input.read(&mut buffer)?;
        if read == 0 {
            break;
        }
        output.write_all(&buffer[..read])?;
        progress.add_counts(
            ItemCounts {
                files: 0,
                directories: 0,
                bytes: read as u64,
            },
            "Trwa kopiowanie danych.",
            Some(source),
            Some(destination),
        );
    }

    progress.add_counts(
        ItemCounts {
            files: 1,
            directories: 0,
            bytes: 0,
        },
        "Plik skopiowany.",
        Some(source),
        Some(destination),
    );
    Ok(OperationResult::Done)
}

fn remote_metadata_from_local(metadata: &fs::Metadata) -> RemoteMetadata {
    let mut remote = RemoteMetadata::default()
        .file_type(RemoteFileType::File)
        .size(metadata.len());
    if let Ok(modified) = metadata.modified() {
        remote = remote.modified(modified);
    }
    remote
}

fn is_canceled_io_error(error: &io::Error) -> bool {
    error.kind() == io::ErrorKind::Interrupted
}

fn remote_temp_dir(name: &str) -> io::Result<PathBuf> {
    let path = std::env::temp_dir().join(name);
    fs::create_dir_all(&path)?;
    Ok(path)
}

fn unique_temp_file_path(prefix: &str, display_name: &str) -> io::Result<PathBuf> {
    let temp_dir = remote_temp_dir(prefix)?;
    let stamp = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_nanos();
    let file_name = display_name.trim();
    let sanitized = if file_name.is_empty() {
        "element".to_string()
    } else {
        sanitize_file_component(file_name)
    };
    Ok(temp_dir.join(format!("{stamp}_{sanitized}")))
}

fn unique_staging_destination(
    staging_root: &Path,
    target: &Path,
    used_names: &mut HashSet<String>,
) -> PathBuf {
    let original_name = target
        .file_name()
        .map(|name| name.to_string_lossy().to_string())
        .filter(|name| !name.is_empty())
        .unwrap_or_else(|| "element".to_string());
    let candidate = sanitize_file_component(&original_name);
    if used_names.insert(candidate.clone()) {
        return staging_root.join(candidate);
    }

    let stem = Path::new(&candidate)
        .file_stem()
        .map(|stem| stem.to_string_lossy().to_string())
        .filter(|stem| !stem.is_empty())
        .unwrap_or_else(|| "element".to_string());
    let extension = Path::new(&candidate)
        .extension()
        .map(|ext| ext.to_string_lossy().to_string());
    for index in 2.. {
        let name = if let Some(extension) = &extension {
            format!("{stem}-{index}.{extension}")
        } else {
            format!("{stem}-{index}")
        };
        if used_names.insert(name.clone()) {
            return staging_root.join(name);
        }
    }
    staging_root.join(candidate)
}

fn current_unix_timestamp() -> u64 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs()
}

fn copy_local_item_plain(source: &Path, destination: &Path) -> io::Result<()> {
    let metadata = fs::symlink_metadata(source)?;
    if metadata.file_type().is_symlink() || metadata.is_file() {
        if let Some(parent) = destination.parent() {
            fs::create_dir_all(parent)?;
        }
        fs::copy(source, destination)?;
        return Ok(());
    }

    if metadata.is_dir() {
        fs::create_dir_all(destination)?;
        for child in fs::read_dir(source)? {
            let child = child?;
            let child_source = child.path();
            let child_destination = destination.join(child.file_name());
            copy_local_item_plain(&child_source, &child_destination)?;
        }
        return Ok(());
    }

    Err(io::Error::other("nieobsługiwany typ elementu"))
}

fn move_local_item_plain(source: &Path, destination: &Path) -> io::Result<()> {
    if destination.exists() {
        remove_existing(destination)?;
    }
    match fs::rename(source, destination) {
        Ok(()) => Ok(()),
        Err(_) => {
            copy_local_item_plain(source, destination)?;
            delete_path_after_move(source)
        }
    }
}

fn execute_elevated_local_operation(operation: ElevatedLocalOperation) -> io::Result<()> {
    match operation {
        ElevatedLocalOperation::Rename {
            source,
            destination,
        } => fs::rename(source, destination),
        ElevatedLocalOperation::CreateDir { path } => fs::create_dir_all(path),
        ElevatedLocalOperation::DeleteTargets { targets } => {
            for target in targets {
                if target.exists() {
                    delete_path_after_move(&target)?;
                }
            }
            Ok(())
        }
        ElevatedLocalOperation::CopyMove {
            sources,
            destination_dir,
            move_mode,
        } => {
            for source in sources {
                let Some(file_name) = source.file_name() else {
                    continue;
                };
                let destination = destination_dir.join(file_name);
                if move_mode {
                    move_local_item_plain(&source, &destination)?;
                } else {
                    if destination.exists() {
                        remove_existing(&destination)?;
                    }
                    copy_local_item_plain(&source, &destination)?;
                }
            }
            Ok(())
        }
    }
}

fn maybe_handle_elevated_local_request() -> Option<i32> {
    let mut args = std::env::args_os();
    let _ = args.next();
    let flag = args.next()?;
    if flag != "--elevated-op" {
        return None;
    }
    let path = PathBuf::from(args.next()?);
    let mut request = match fs::read_to_string(&path)
        .ok()
        .and_then(|text| serde_json::from_str::<ElevatedLocalRequestFile>(&text).ok())
    {
        Some(request) => request,
        None => return Some(1),
    };

    match execute_elevated_local_operation(request.operation.clone()) {
        Ok(()) => {
            request.success = Some(true);
            request.error = None;
        }
        Err(error) => {
            request.success = Some(false);
            request.error = Some(error.to_string());
        }
    }

    let _ = serde_json::to_string_pretty(&request)
        .map_err(|error| io::Error::other(error.to_string()))
        .and_then(|text| fs::write(&path, text));
    Some(if request.success == Some(true) { 0 } else { 1 })
}

unsafe fn run_elevated_local_operation(operation: ElevatedLocalOperation) -> io::Result<()> {
    let request_path = unique_temp_file_path("AmigaFmNativeElevated", "request.json")?;
    let request = ElevatedLocalRequestFile {
        operation,
        success: None,
        error: None,
    };
    let contents = serde_json::to_string_pretty(&request)
        .map_err(|error| io::Error::other(error.to_string()))?;
    fs::write(&request_path, contents)?;

    let executable = std::env::current_exe()?;
    let verb = wide("runas");
    let file = path_wide(&executable);
    let parameters = wide(&format!("--elevated-op \"{}\"", request_path.display()));
    let mut execute_info = SHELLEXECUTEINFOW {
        cbSize: std::mem::size_of::<SHELLEXECUTEINFOW>() as u32,
        fMask: SEE_MASK_NOCLOSEPROCESS,
        hwnd: null_mut(),
        lpVerb: verb.as_ptr(),
        lpFile: file.as_ptr(),
        lpParameters: parameters.as_ptr(),
        lpDirectory: null(),
        nShow: SW_SHOW,
        hInstApp: null_mut(),
        ..unsafe { std::mem::zeroed() }
    };

    if ShellExecuteExW(&mut execute_info) == 0 {
        let error = io::Error::last_os_error();
        let _ = fs::remove_file(&request_path);
        return Err(error);
    }

    if !execute_info.hProcess.is_null() {
        WaitForSingleObject(execute_info.hProcess, INFINITE);
        let mut exit_code = 1u32;
        let _ = GetExitCodeProcess(execute_info.hProcess, &mut exit_code);
        CloseHandle(execute_info.hProcess);
    }

    let result = fs::read_to_string(&request_path)
        .ok()
        .and_then(|text| serde_json::from_str::<ElevatedLocalRequestFile>(&text).ok());
    let _ = fs::remove_file(&request_path);

    match result {
        Some(result) if result.success == Some(true) => Ok(()),
        Some(result) if result.success == Some(false) => {
            Err(io::Error::other(result.error.unwrap_or_else(|| {
                "operacja podniesiona nie powiodła się".to_string()
            })))
        }
        _ => Err(io::Error::other(
            "nie udało się odczytać wyniku operacji z podniesionymi uprawnieniami",
        )),
    }
}

fn summarize_remote_path(remote: &RemoteLocation, path: &Path) -> io::Result<ItemCounts> {
    if is_sftp_sudo_location(remote) {
        let (is_dir, size) = stat_sftp_path_via_sudo(remote, path)?;
        if !is_dir {
            return Ok(ItemCounts {
                files: 1,
                directories: 0,
                bytes: size,
            });
        }
        let mut counts = ItemCounts {
            files: 0,
            directories: 1,
            bytes: 0,
        };
        for (_, child_path, _, _) in list_sftp_children_via_sudo(remote, path)? {
            counts.add(summarize_remote_path(remote, &child_path)?);
        }
        return Ok(counts);
    }
    let mut client = connect_remote_client(&remote.resource)?;
    let result = summarize_remote_path_with_client(&mut client, path);
    client.disconnect();
    result
}

fn summarize_remote_path_with_client(
    client: &mut RemoteClient,
    path: &Path,
) -> io::Result<ItemCounts> {
    let entry = client.stat(path)?;
    if entry.is_dir() {
        let mut counts = ItemCounts {
            files: 0,
            directories: 1,
            bytes: 0,
        };
        for child in client.list_dir(path)? {
            counts.add(summarize_remote_path_with_client(client, child.path())?);
        }
        Ok(counts)
    } else {
        Ok(ItemCounts {
            files: 1,
            directories: 0,
            bytes: entry.metadata().size,
        })
    }
}

fn delete_remote_path_after_move_with_client(
    client: &mut RemoteClient,
    path: &Path,
) -> io::Result<()> {
    delete_remote_path_recursive_with_client(client, path)
}

fn delete_remote_path_recursive_with_client(
    client: &mut RemoteClient,
    path: &Path,
) -> io::Result<()> {
    let entry = match client.stat(path) {
        Ok(entry) => entry,
        Err(error) if is_not_found_error(&error) => return Ok(()),
        Err(error) => return Err(error),
    };

    if entry.is_dir() {
        let children = match client.list_dir(path) {
            Ok(children) => children,
            Err(error) if is_not_found_error(&error) => return Ok(()),
            Err(error) => return Err(error),
        };
        for child in children {
            let name = child.name();
            if name == "." || name == ".." {
                continue;
            }
            delete_remote_path_recursive_with_client(client, child.path())?;
        }
        match client.remove_dir(path) {
            Ok(()) => Ok(()),
            Err(error) if is_not_found_error(&error) => Ok(()),
            Err(error) => Err(error),
        }
    } else {
        match client.remove_file(path) {
            Ok(()) => Ok(()),
            Err(error) if is_not_found_error(&error) => Ok(()),
            Err(error) => Err(error),
        }
    }
}

fn remove_existing_remote_with_client(client: &mut RemoteClient, path: &Path) -> io::Result<()> {
    if client.exists(path)? {
        delete_remote_path_after_move_with_client(client, path)?;
    }
    Ok(())
}

fn remote_exists_via_sudo(remote: &RemoteLocation, path: &Path) -> io::Result<bool> {
    let output = exec_sftp_sudo_command(
        &remote.resource,
        &format!(
            "if [ -e {} ]; then printf '1'; else printf '0'; fi",
            shell_quote_posix(&remote_shell_path(path))
        ),
    )?;
    Ok(output.trim() == "1")
}

fn remove_existing_remote_via_sudo(remote: &RemoteLocation, path: &Path) -> io::Result<()> {
    if remote_exists_via_sudo(remote, path)? {
        exec_sftp_sudo_command(
            &remote.resource,
            &format!("rm -rf -- {}", shell_quote_posix(&remote_shell_path(path))),
        )?;
    }
    Ok(())
}

unsafe fn copy_local_to_remote_via_sudo(
    remote: &RemoteLocation,
    source: &Path,
    destination: &Path,
    summary: ItemCounts,
    progress: &mut ProgressReporter,
) -> io::Result<OperationResult> {
    if progress.is_canceled() {
        return Ok(OperationResult::Canceled);
    }
    if remote_exists_via_sudo(remote, destination)? {
        match progress.ask_conflict(destination)? {
            ConflictChoice::Replace => remove_existing_remote_via_sudo(remote, destination)?,
            ConflictChoice::Skip => return Ok(OperationResult::Skipped),
            ConflictChoice::Cancel => return Ok(OperationResult::Canceled),
        }
    }

    let metadata = fs::symlink_metadata(source)?;
    if metadata.file_type().is_symlink() || metadata.is_file() {
        progress.update("Kopiowanie pliku.", Some(source), Some(destination));
        let temp_remote = exec_remote_command(&remote.resource, "mktemp")
            .map(|(_, output)| output.trim().to_string())
            .and_then(|output| {
                if output.is_empty() {
                    Err(io::Error::other(
                        "nie udało się przygotować pliku tymczasowego na serwerze",
                    ))
                } else {
                    Ok(PathBuf::from(output))
                }
            })?;
        let mut client = connect_remote_client(&remote.resource)?;
        let result = client.create_file(
            &temp_remote,
            &remote_metadata_from_local(&metadata),
            Box::new(fs::File::open(source)?),
        );
        client.disconnect();
        match result {
            Ok(bytes) => {
                let parent = destination
                    .parent()
                    .filter(|parent| !parent.as_os_str().is_empty())
                    .unwrap_or_else(|| Path::new("/"));
                exec_sftp_sudo_command(
                    &remote.resource,
                    &format!(
                        "mkdir -p -- {} && cp -- {} {} && rm -f -- {}",
                        shell_quote_posix(&remote_shell_path(parent)),
                        shell_quote_posix(&remote_shell_path(&temp_remote)),
                        shell_quote_posix(&remote_shell_path(destination)),
                        shell_quote_posix(&remote_shell_path(&temp_remote))
                    ),
                )?;
                progress.add_counts(
                    ItemCounts {
                        files: 1,
                        directories: 0,
                        bytes,
                    },
                    "Plik skopiowany.",
                    Some(source),
                    Some(destination),
                );
                return Ok(OperationResult::Done);
            }
            Err(error) if is_canceled_io_error(&error) || progress.is_canceled() => {
                let _ = exec_remote_command(
                    &remote.resource,
                    &format!(
                        "rm -f -- {}",
                        shell_quote_posix(&remote_shell_path(&temp_remote))
                    ),
                );
                return Ok(OperationResult::Canceled);
            }
            Err(error) => {
                let _ = exec_remote_command(
                    &remote.resource,
                    &format!(
                        "rm -f -- {}",
                        shell_quote_posix(&remote_shell_path(&temp_remote))
                    ),
                );
                return Err(error);
            }
        }
    }

    if metadata.is_dir() {
        progress.update("Tworzenie katalogu.", Some(source), Some(destination));
        exec_sftp_sudo_command(
            &remote.resource,
            &format!(
                "mkdir -p -- {}",
                shell_quote_posix(&remote_shell_path(destination))
            ),
        )?;
        progress.add_counts(
            ItemCounts {
                files: 0,
                directories: 1,
                bytes: 0,
            },
            "Katalog utworzony.",
            Some(source),
            Some(destination),
        );
        for child in fs::read_dir(source)? {
            if progress.is_canceled() {
                return Ok(OperationResult::Canceled);
            }
            let child = child?;
            let child_source = child.path();
            let child_destination = destination.join(child.file_name());
            let child_summary = summarize_path(&child_source)?;
            match copy_local_to_remote_via_sudo(
                remote,
                &child_source,
                &child_destination,
                child_summary,
                progress,
            )? {
                OperationResult::Done | OperationResult::Skipped => {}
                OperationResult::Canceled => return Ok(OperationResult::Canceled),
            }
        }
        return Ok(OperationResult::Done);
    }

    progress.add_counts(
        summary,
        "Element zakończony.",
        Some(source),
        Some(destination),
    );
    Ok(OperationResult::Done)
}

unsafe fn copy_remote_to_local_via_sudo(
    remote: &RemoteLocation,
    source: &Path,
    destination: &Path,
    _summary: ItemCounts,
    progress: &mut ProgressReporter,
) -> io::Result<OperationResult> {
    if progress.is_canceled() {
        return Ok(OperationResult::Canceled);
    }
    if destination.exists() {
        match progress.ask_conflict(destination)? {
            ConflictChoice::Replace => remove_existing(destination)?,
            ConflictChoice::Skip => return Ok(OperationResult::Skipped),
            ConflictChoice::Cancel => return Ok(OperationResult::Canceled),
        }
    }

    let (is_dir, _size) = stat_sftp_path_via_sudo(remote, source)?;
    if is_dir {
        progress.update("Tworzenie katalogu.", Some(source), Some(destination));
        fs::create_dir_all(destination)?;
        progress.add_counts(
            ItemCounts {
                files: 0,
                directories: 1,
                bytes: 0,
            },
            "Katalog utworzony.",
            Some(source),
            Some(destination),
        );
        for (name, child_source, _, _) in list_sftp_children_via_sudo(remote, source)? {
            if progress.is_canceled() {
                return Ok(OperationResult::Canceled);
            }
            let child_destination = destination.join(name);
            let child_summary = summarize_remote_path(remote, &child_source)?;
            match copy_remote_to_local_via_sudo(
                remote,
                &child_source,
                &child_destination,
                child_summary,
                progress,
            )? {
                OperationResult::Done | OperationResult::Skipped => {}
                OperationResult::Canceled => return Ok(OperationResult::Canceled),
            }
        }
        return Ok(OperationResult::Done);
    }

    progress.update("Kopiowanie pliku.", Some(source), Some(destination));
    let temp_path = download_sftp_file_via_sudo_to_temp(remote, source, progress)?;
    if let Some(parent) = destination.parent() {
        fs::create_dir_all(parent)?;
    }
    fs::copy(&temp_path, destination)?;
    let _ = fs::remove_file(&temp_path);
    Ok(OperationResult::Done)
}

unsafe fn copy_local_to_remote_with_sftp_fallback(
    remote: &RemoteLocation,
    fallback: Option<&RemoteLocation>,
    source: &Path,
    destination: &Path,
    summary: ItemCounts,
    progress: &mut ProgressReporter,
) -> io::Result<OperationResult> {
    if remote.uses_sftp_sudo() {
        return copy_local_to_remote_via_sudo(remote, source, destination, summary, progress);
    }

    let mut client = connect_remote_client(&remote.resource)?;
    let result =
        copy_local_to_remote_with_client(&mut client, source, destination, summary, progress);
    client.disconnect();
    match result {
        Err(error) if is_sftp_retry_error(&error) => {
            if let Some(fallback) = fallback {
                progress.update(
                    "Element wymaga sudo, ponawianie tylko dla tego elementu.",
                    Some(source),
                    Some(destination),
                );
                copy_local_to_remote_via_sudo(fallback, source, destination, summary, progress)
            } else {
                Err(error)
            }
        }
        other => other,
    }
}

unsafe fn copy_remote_to_local_with_sftp_fallback(
    remote: &RemoteLocation,
    fallback: Option<&RemoteLocation>,
    source: &Path,
    destination: &Path,
    summary: ItemCounts,
    progress: &mut ProgressReporter,
) -> io::Result<OperationResult> {
    if remote.uses_sftp_sudo() {
        return copy_remote_to_local_via_sudo(remote, source, destination, summary, progress);
    }

    let mut client = connect_remote_client(&remote.resource)?;
    let result =
        copy_remote_to_local_with_client(&mut client, source, destination, summary, progress);
    client.disconnect();
    match result {
        Err(error) if is_sftp_retry_error(&error) => {
            if let Some(fallback) = fallback {
                progress.update(
                    "Element wymaga sudo, ponawianie tylko dla tego elementu.",
                    Some(source),
                    Some(destination),
                );
                copy_remote_to_local_via_sudo(fallback, source, destination, summary, progress)
            } else {
                Err(error)
            }
        }
        other => other,
    }
}

unsafe fn move_remote_with_sftp_fallback(
    remote: &RemoteLocation,
    fallback: Option<&RemoteLocation>,
    source: &Path,
    destination: &Path,
    summary: ItemCounts,
    progress: &mut ProgressReporter,
) -> io::Result<OperationResult> {
    if remote.uses_sftp_sudo() {
        progress.update("Przenoszenie elementu.", Some(source), Some(destination));
        remote_rename(remote, source, destination)?;
        progress.add_counts(
            summary,
            "Element przeniesiony.",
            Some(source),
            Some(destination),
        );
        return Ok(OperationResult::Done);
    }

    let mut client = connect_remote_client(&remote.resource)?;
    let result = move_remote_with_client(&mut client, source, destination, summary, progress);
    client.disconnect();
    match result {
        Err(error) if is_sftp_retry_error(&error) => {
            if let Some(fallback) = fallback {
                progress.update(
                    "Element wymaga sudo, przenoszenie tylko tego elementu.",
                    Some(source),
                    Some(destination),
                );
                remote_rename(fallback, source, destination)?;
                progress.add_counts(
                    summary,
                    "Element przeniesiony.",
                    Some(source),
                    Some(destination),
                );
                Ok(OperationResult::Done)
            } else {
                Err(error)
            }
        }
        other => other,
    }
}

unsafe fn delete_remote_path_after_move_with_sftp_fallback(
    remote: &RemoteLocation,
    fallback: Option<&RemoteLocation>,
    path: &Path,
) -> io::Result<()> {
    if remote.uses_sftp_sudo() {
        return remote_delete_targets(remote, &[path.to_path_buf()], &[]);
    }

    let mut client = connect_remote_client(&remote.resource)?;
    let result = delete_remote_path_after_move_with_client(&mut client, path);
    client.disconnect();
    match result {
        Err(error) if is_sftp_retry_error(&error) => {
            if let Some(fallback) = fallback {
                remote_delete_targets(fallback, &[path.to_path_buf()], &[])
            } else {
                Err(error)
            }
        }
        other => other,
    }
}

unsafe fn copy_remote_to_remote_with_sftp_fallback(
    source_remote: &RemoteLocation,
    destination_remote: &RemoteLocation,
    source_fallback: Option<&RemoteLocation>,
    destination_fallback: Option<&RemoteLocation>,
    source: &Path,
    destination: &Path,
    summary: ItemCounts,
    progress: &mut ProgressReporter,
) -> io::Result<OperationResult> {
    if source_remote.resource.protocol == NetworkProtocol::Sftp
        || destination_remote.resource.protocol == NetworkProtocol::Sftp
        || source_remote.uses_sftp_sudo()
        || destination_remote.uses_sftp_sudo()
        || source_fallback.is_some()
        || destination_fallback.is_some()
    {
        return copy_remote_to_remote_via_temp_with_fallbacks(
            source_remote,
            destination_remote,
            source_fallback.or(Some(source_remote)),
            destination_fallback.or(Some(destination_remote)),
            source,
            destination,
            summary,
            progress,
        );
    }

    let mut source_client = connect_remote_client(&source_remote.resource)?;
    let mut destination_client = connect_remote_client(&destination_remote.resource)?;
    let result = copy_remote_to_remote_with_clients(
        &mut source_client,
        &mut destination_client,
        source,
        destination,
        summary,
        progress,
    );
    destination_client.disconnect();
    source_client.disconnect();
    match result {
        Err(error)
            if is_sftp_retry_error(&error)
                && (source_fallback.is_some() || destination_fallback.is_some()) =>
        {
            copy_remote_to_remote_via_temp_with_fallbacks(
                source_remote,
                destination_remote,
                source_fallback,
                destination_fallback,
                source,
                destination,
                summary,
                progress,
            )
        }
        other => other,
    }
}

unsafe fn copy_remote_to_remote_via_temp_with_fallbacks(
    source_remote: &RemoteLocation,
    destination_remote: &RemoteLocation,
    source_fallback: Option<&RemoteLocation>,
    destination_fallback: Option<&RemoteLocation>,
    source: &Path,
    destination: &Path,
    summary: ItemCounts,
    progress: &mut ProgressReporter,
) -> io::Result<OperationResult> {
    let file_name = source
        .file_name()
        .map(|name| name.to_string_lossy().to_string())
        .unwrap_or_else(|| "remote-item".to_string());
    let temp_path = unique_temp_file_path("AmigaFmNativeRemoteBridge", &file_name)?;
    let download_result = copy_remote_to_local_with_sftp_fallback(
        source_remote,
        source_fallback,
        source,
        &temp_path,
        summary,
        progress,
    );
    match download_result {
        Ok(OperationResult::Done) | Ok(OperationResult::Skipped) => {}
        Ok(OperationResult::Canceled) => {
            let _ = remove_existing(&temp_path);
            return Ok(OperationResult::Canceled);
        }
        Err(error) => {
            let _ = remove_existing(&temp_path);
            return Err(io::Error::new(
                error.kind(),
                format!("źródło SFTP: {error}"),
            ));
        }
    }

    let upload_summary = summarize_path(&temp_path).unwrap_or(summary);
    let upload_result = copy_local_to_remote_with_sftp_fallback(
        destination_remote,
        destination_fallback,
        &temp_path,
        destination,
        upload_summary,
        progress,
    );
    let _ = remove_existing(&temp_path);
    upload_result.map_err(|error| io::Error::new(error.kind(), format!("cel SFTP: {error}")))
}

unsafe fn move_remote_with_client(
    client: &mut RemoteClient,
    source: &Path,
    destination: &Path,
    summary: ItemCounts,
    progress: &mut ProgressReporter,
) -> io::Result<OperationResult> {
    if progress.is_canceled() {
        return Ok(OperationResult::Canceled);
    }
    if client.exists(destination)? {
        match progress.ask_conflict(destination)? {
            ConflictChoice::Replace => remove_existing_remote_with_client(client, destination)?,
            ConflictChoice::Skip => return Ok(OperationResult::Skipped),
            ConflictChoice::Cancel => return Ok(OperationResult::Canceled),
        }
    }

    progress.update("Przenoszenie elementu.", Some(source), Some(destination));
    client.rename(source, destination)?;
    progress.add_counts(
        summary,
        "Element przeniesiony.",
        Some(source),
        Some(destination),
    );
    Ok(OperationResult::Done)
}

unsafe fn copy_local_to_remote_with_client(
    client: &mut RemoteClient,
    source: &Path,
    destination: &Path,
    summary: ItemCounts,
    progress: &mut ProgressReporter,
) -> io::Result<OperationResult> {
    if progress.is_canceled() {
        return Ok(OperationResult::Canceled);
    }
    if client.exists(destination)? {
        match progress.ask_conflict(destination)? {
            ConflictChoice::Replace => remove_existing_remote_with_client(client, destination)?,
            ConflictChoice::Skip => return Ok(OperationResult::Skipped),
            ConflictChoice::Cancel => return Ok(OperationResult::Canceled),
        }
    }

    let metadata = fs::symlink_metadata(source)?;
    if metadata.file_type().is_symlink() || metadata.is_file() {
        progress.update("Kopiowanie pliku.", Some(source), Some(destination));
        let result = client.create_file(
            destination,
            &remote_metadata_from_local(&metadata),
            Box::new(fs::File::open(source)?),
        );
        match result {
            Ok(bytes) => {
                progress.add_counts(
                    ItemCounts {
                        files: 1,
                        directories: 0,
                        bytes,
                    },
                    "Plik skopiowany.",
                    Some(source),
                    Some(destination),
                );
            }
            Err(error) if is_canceled_io_error(&error) || progress.is_canceled() => {
                let _ = client.remove_file(destination);
                return Ok(OperationResult::Canceled);
            }
            Err(error) => {
                let _ = client.remove_file(destination);
                return Err(error);
            }
        }
        return Ok(OperationResult::Done);
    }

    if metadata.is_dir() {
        progress.update("Tworzenie katalogu.", Some(source), Some(destination));
        client.create_dir(destination)?;
        progress.add_counts(
            ItemCounts {
                files: 0,
                directories: 1,
                bytes: 0,
            },
            "Katalog utworzony.",
            Some(source),
            Some(destination),
        );
        for child in fs::read_dir(source)? {
            if progress.is_canceled() {
                return Ok(OperationResult::Canceled);
            }
            let child = child?;
            let child_source = child.path();
            let child_destination = destination.join(child.file_name());
            let child_summary = summarize_path(&child_source)?;
            match copy_local_to_remote_with_client(
                client,
                &child_source,
                &child_destination,
                child_summary,
                progress,
            )? {
                OperationResult::Done | OperationResult::Skipped => {}
                OperationResult::Canceled => return Ok(OperationResult::Canceled),
            }
        }
        return Ok(OperationResult::Done);
    }

    progress.add_counts(
        summary,
        "Element zakończony.",
        Some(source),
        Some(destination),
    );
    Ok(OperationResult::Done)
}

unsafe fn copy_remote_to_local_with_client(
    client: &mut RemoteClient,
    source: &Path,
    destination: &Path,
    _summary: ItemCounts,
    progress: &mut ProgressReporter,
) -> io::Result<OperationResult> {
    if progress.is_canceled() {
        return Ok(OperationResult::Canceled);
    }
    if destination.exists() {
        match progress.ask_conflict(destination)? {
            ConflictChoice::Replace => remove_existing(destination)?,
            ConflictChoice::Skip => return Ok(OperationResult::Skipped),
            ConflictChoice::Cancel => return Ok(OperationResult::Canceled),
        }
    }

    let entry = client.stat(source)?;
    if entry.is_dir() {
        progress.update("Tworzenie katalogu.", Some(source), Some(destination));
        fs::create_dir_all(destination)?;
        progress.add_counts(
            ItemCounts {
                files: 0,
                directories: 1,
                bytes: 0,
            },
            "Katalog utworzony.",
            Some(source),
            Some(destination),
        );
        for child in client.list_dir(source)? {
            if progress.is_canceled() {
                return Ok(OperationResult::Canceled);
            }
            let child_destination = destination.join(child.name());
            let child_summary = summarize_remote_path_with_client(client, child.path())?;
            match copy_remote_to_local_with_client(
                client,
                child.path(),
                &child_destination,
                child_summary,
                progress,
            )? {
                OperationResult::Done | OperationResult::Skipped => {}
                OperationResult::Canceled => return Ok(OperationResult::Canceled),
            }
        }
        return Ok(OperationResult::Done);
    }

    progress.update("Kopiowanie pliku.", Some(source), Some(destination));
    let result = client.download_to(source, Box::new(fs::File::create(destination)?));
    match result {
        Ok(bytes) => {
            progress.add_counts(
                ItemCounts {
                    files: 1,
                    directories: 0,
                    bytes,
                },
                "Plik skopiowany.",
                Some(source),
                Some(destination),
            );
        }
        Err(error) if is_canceled_io_error(&error) || progress.is_canceled() => {
            let _ = fs::remove_file(destination);
            return Ok(OperationResult::Canceled);
        }
        Err(error) => {
            let _ = fs::remove_file(destination);
            return Err(error);
        }
    }
    Ok(OperationResult::Done)
}

unsafe fn copy_remote_to_remote_with_clients(
    source_client: &mut RemoteClient,
    destination_client: &mut RemoteClient,
    source: &Path,
    destination: &Path,
    _summary: ItemCounts,
    progress: &mut ProgressReporter,
) -> io::Result<OperationResult> {
    if progress.is_canceled() {
        return Ok(OperationResult::Canceled);
    }
    if destination_client.exists(destination)? {
        match progress.ask_conflict(destination)? {
            ConflictChoice::Replace => {
                remove_existing_remote_with_client(destination_client, destination)?
            }
            ConflictChoice::Skip => return Ok(OperationResult::Skipped),
            ConflictChoice::Cancel => return Ok(OperationResult::Canceled),
        }
    }

    let entry = source_client.stat(source)?;
    if entry.is_dir() {
        progress.update("Tworzenie katalogu.", Some(source), Some(destination));
        destination_client.create_dir(destination)?;
        progress.add_counts(
            ItemCounts {
                files: 0,
                directories: 1,
                bytes: 0,
            },
            "Katalog utworzony.",
            Some(source),
            Some(destination),
        );
        for child in source_client.list_dir(source)? {
            if progress.is_canceled() {
                return Ok(OperationResult::Canceled);
            }
            let child_source = child.path().to_path_buf();
            let child_destination = destination.join(child.name());
            let child_summary = summarize_remote_path_with_client(source_client, &child_source)?;
            match copy_remote_to_remote_with_clients(
                source_client,
                destination_client,
                &child_source,
                &child_destination,
                child_summary,
                progress,
            )? {
                OperationResult::Done | OperationResult::Skipped => {}
                OperationResult::Canceled => return Ok(OperationResult::Canceled),
            }
        }
        return Ok(OperationResult::Done);
    }

    progress.update(
        "Pobieranie pliku z zasobu źródłowego.",
        Some(source),
        Some(destination),
    );
    let temp_path = unique_temp_file_path("AmigaFmNativeTransfers", &entry.name())?;
    {
        let result = source_client.download_to(source, Box::new(fs::File::create(&temp_path)?));
        if let Err(error) = result {
            let _ = fs::remove_file(&temp_path);
            if is_canceled_io_error(&error) || progress.is_canceled() {
                return Ok(OperationResult::Canceled);
            }
            return Err(error);
        }
    }

    progress.update(
        "Wysyłanie pliku do zasobu docelowego.",
        Some(source),
        Some(destination),
    );
    let upload_result = {
        let metadata = RemoteMetadata::default()
            .file_type(RemoteFileType::File)
            .size(entry.metadata().size);
        destination_client.create_file(
            destination,
            &metadata,
            Box::new(fs::File::open(&temp_path)?),
        )
    };
    let _ = fs::remove_file(&temp_path);
    match upload_result {
        Ok(bytes) => {
            progress.add_counts(
                ItemCounts {
                    files: 1,
                    directories: 0,
                    bytes,
                },
                "Plik skopiowany.",
                Some(source),
                Some(destination),
            );
            Ok(OperationResult::Done)
        }
        Err(error) if is_canceled_io_error(&error) || progress.is_canceled() => {
            let _ = destination_client.remove_file(destination);
            Ok(OperationResult::Canceled)
        }
        Err(error) => {
            let _ = destination_client.remove_file(destination);
            Err(error)
        }
    }
}

unsafe fn delete_path_with_progress(
    path: &Path,
    progress: &mut ProgressReporter,
) -> io::Result<OperationResult> {
    if progress.is_canceled() {
        return Ok(OperationResult::Canceled);
    }
    let metadata = fs::symlink_metadata(path)?;
    if metadata.file_type().is_symlink() || metadata.is_file() {
        progress.update("Usuwanie pliku.", Some(path), None);
        fs::remove_file(path)?;
        progress.add_counts(
            ItemCounts {
                files: 1,
                directories: 0,
                bytes: metadata.len(),
            },
            "Plik usunięty.",
            Some(path),
            None,
        );
        Ok(OperationResult::Done)
    } else if metadata.is_dir() {
        for child in fs::read_dir(path)? {
            if progress.is_canceled() {
                return Ok(OperationResult::Canceled);
            }
            let child = child?;
            match delete_path_with_progress(&child.path(), progress)? {
                OperationResult::Done | OperationResult::Skipped => {}
                OperationResult::Canceled => return Ok(OperationResult::Canceled),
            }
        }
        progress.update("Usuwanie katalogu.", Some(path), None);
        fs::remove_dir(path)?;
        progress.add_counts(
            ItemCounts {
                files: 0,
                directories: 1,
                bytes: 0,
            },
            "Katalog usunięty.",
            Some(path),
            None,
        );
        Ok(OperationResult::Done)
    } else {
        progress.update("Usuwanie elementu.", Some(path), None);
        fs::remove_file(path)?;
        Ok(OperationResult::Done)
    }
}

fn delete_path_after_move(path: &Path) -> io::Result<()> {
    let metadata = fs::symlink_metadata(path)?;
    if metadata.file_type().is_symlink() || metadata.is_file() {
        fs::remove_file(path)
    } else if metadata.is_dir() {
        fs::remove_dir_all(path)
    } else {
        fs::remove_file(path)
    }
}

fn remove_existing(path: &Path) -> io::Result<()> {
    if path.exists() {
        delete_path_after_move(path)?;
    }
    Ok(())
}

fn validate_path_removed(path: &Path) -> io::Result<()> {
    if path.exists() {
        Err(io::Error::other(format!(
            "element nadal istnieje po usunięciu: {}",
            path.display()
        )))
    } else {
        Ok(())
    }
}

fn validate_destination(source: &Path, destination_dir: &Path) -> Result<(), String> {
    let Some(name) = source.file_name() else {
        return Err("nie można określić nazwy źródła".to_string());
    };
    let destination = destination_dir.join(name);
    if source == destination {
        return Err("źródło i cel są takie same".to_string());
    }
    if source.is_dir() && destination_dir.starts_with(source) {
        return Err("nie można kopiować lub przenosić katalogu do jego wnętrza".to_string());
    }
    Ok(())
}

fn open_with_system(path: &Path) -> io::Result<()> {
    open_with_system_target(&path.display().to_string())
}

fn open_with_system_target(target: &str) -> io::Result<()> {
    let operation = wide("open");
    let target_wide = wide(target);
    let result = unsafe {
        ShellExecuteW(
            null_mut(),
            operation.as_ptr(),
            target_wide.as_ptr(),
            null(),
            null(),
            SW_SHOW,
        )
    } as isize;

    if result > 32 {
        return Ok(());
    }

    if !target.contains("://") && !target.starts_with("\\\\") {
        let status = Command::new("rundll32.exe")
            .arg("shell32.dll,OpenAs_RunDLL")
            .arg(target)
            .creation_flags(CREATE_NO_WINDOW_FLAG)
            .status()?;

        if status.success() {
            return Ok(());
        }
    }

    Err(io::Error::other("system nie otworzył zasobu"))
}

fn confirm_exit(hwnd: HWND) -> bool {
    unsafe {
        let nvda = app_state_mut(hwnd).map(|app| &app.nvda);
        show_choice_prompt(
            hwnd,
            "Zamknij program",
            vec!["Czy zamknąć program?".to_string()],
            vec![
                DialogButton {
                    id: ID_DIALOG_YES,
                    label: "Tak",
                    is_default: true,
                },
                DialogButton {
                    id: ID_DIALOG_NO,
                    label: "Nie",
                    is_default: false,
                },
            ],
            nvda,
        ) == ID_DIALOG_YES
    }
}

unsafe fn register_class(
    class_name: &str,
    window_proc: Option<unsafe extern "system" fn(HWND, u32, WPARAM, LPARAM) -> LRESULT>,
) -> Result<(), String> {
    let instance = GetModuleHandleW(null());
    let class_name_w = wide(class_name);
    let icon = LoadIconW(null_mut(), IDI_APPLICATION);
    let class = WNDCLASSEXW {
        cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
        style: CS_HREDRAW | CS_VREDRAW,
        lpfnWndProc: window_proc,
        hInstance: instance,
        hIcon: icon,
        hCursor: LoadCursorW(null_mut(), IDC_ARROW),
        hbrBackground: GetStockObject(COLOR_WINDOWFRAME) as HBRUSH,
        lpszClassName: class_name_w.as_ptr(),
        cbClsExtra: 0,
        cbWndExtra: 0,
        lpszMenuName: null(),
        hIconSm: icon,
    };

    let result = RegisterClassExW(&class);
    if result == 0 {
        let error = GetLastError();
        if error != 1410 {
            return Err(last_error_message("nie udało się zarejestrować klasy okna"));
        }
    }
    Ok(())
}

unsafe fn app_state_mut(hwnd: HWND) -> Option<&'static mut AppState> {
    let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut AppState;
    ptr.as_mut()
}

unsafe fn input_dialog_state_mut(hwnd: HWND) -> Option<&'static mut InputDialogState> {
    let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut InputDialogState;
    ptr.as_mut()
}

unsafe fn search_dialog_state_mut(hwnd: HWND) -> Option<&'static mut SearchDialogState> {
    let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut SearchDialogState;
    ptr.as_mut()
}

unsafe fn network_dialog_state_mut(hwnd: HWND) -> Option<&'static mut NetworkDialogState> {
    let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut NetworkDialogState;
    ptr.as_mut()
}

unsafe fn discovery_dialog_state_mut(hwnd: HWND) -> Option<&'static mut DiscoveryDialogState> {
    let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut DiscoveryDialogState;
    ptr.as_mut()
}

unsafe fn archive_create_dialog_state_mut(
    hwnd: HWND,
) -> Option<&'static mut ArchiveCreateDialogState> {
    let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut ArchiveCreateDialogState;
    ptr.as_mut()
}

unsafe fn operation_prompt_state_mut(hwnd: HWND) -> Option<&'static mut OperationPromptState> {
    let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut OperationPromptState;
    ptr.as_mut()
}

unsafe fn progress_dialog_state_mut(hwnd: HWND) -> Option<&'static mut ProgressDialogState> {
    let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut ProgressDialogState;
    ptr.as_mut()
}

fn wide(text: &str) -> Vec<u16> {
    std::ffi::OsStr::new(text)
        .encode_wide()
        .chain(std::iter::once(0))
        .collect()
}

fn path_wide(path: &Path) -> Vec<u16> {
    path.as_os_str()
        .encode_wide()
        .chain(std::iter::once(0))
        .collect()
}

unsafe fn choose_file_dialog(owner: HWND, title: &str, initial: &str) -> Option<PathBuf> {
    let mut file_buffer = vec![0u16; 32768];
    let initial_path = initial.trim();
    if !initial_path.is_empty() {
        let existing = path_wide(Path::new(initial_path));
        let len = existing.len().saturating_sub(1).min(file_buffer.len() - 1);
        file_buffer[..len].copy_from_slice(&existing[..len]);
    }

    let filter = wide("Wszystkie pliki\0*.*\0\0");
    let title = wide(title);
    let initial_dir = Path::new(initial_path)
        .parent()
        .filter(|path| !path.as_os_str().is_empty())
        .map(path_wide);

    let mut dialog: OPENFILENAMEW = std::mem::zeroed();
    dialog.lStructSize = std::mem::size_of::<OPENFILENAMEW>() as u32;
    dialog.hwndOwner = owner;
    dialog.lpstrFilter = filter.as_ptr();
    dialog.lpstrFile = file_buffer.as_mut_ptr();
    dialog.nMaxFile = file_buffer.len() as u32;
    dialog.lpstrInitialDir = initial_dir
        .as_ref()
        .map(|value| value.as_ptr())
        .unwrap_or_else(null);
    dialog.lpstrTitle = title.as_ptr();
    dialog.Flags =
        OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY | OFN_NOCHANGEDIR;

    if GetOpenFileNameW(&mut dialog) == 0 {
        return None;
    }

    let len = file_buffer
        .iter()
        .position(|ch| *ch == 0)
        .unwrap_or(file_buffer.len());
    if len == 0 {
        None
    } else {
        Some(PathBuf::from(String::from_utf16_lossy(&file_buffer[..len])))
    }
}

fn preferred_drop_effect_format() -> u32 {
    unsafe { RegisterClipboardFormatW(wide("Preferred DropEffect").as_ptr()) }
}

unsafe fn write_file_paths_to_clipboard(
    owner: HWND,
    paths: &[PathBuf],
    operation: ClipboardOperation,
) -> io::Result<u32> {
    if paths.is_empty() {
        return Err(io::Error::other("brak elementów do schowka"));
    }

    let mut path_list = Vec::<u16>::new();
    for path in paths {
        let wide_path = path_wide(path);
        path_list.extend_from_slice(&wide_path);
    }
    path_list.push(0);

    let dropfiles_size = std::mem::size_of::<DropFilesHeader>();
    let total_size = dropfiles_size + path_list.len() * std::mem::size_of::<u16>();
    let hdrop = GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, total_size);
    if hdrop.is_null() {
        return Err(io::Error::last_os_error());
    }
    let drop_ptr = GlobalLock(hdrop) as *mut u8;
    if drop_ptr.is_null() {
        GlobalFree(hdrop);
        return Err(io::Error::last_os_error());
    }
    let header = drop_ptr as *mut DropFilesHeader;
    (*header).p_files = dropfiles_size as u32;
    (*header).pt_x = 0;
    (*header).pt_y = 0;
    (*header).f_nc = 0;
    (*header).f_wide = 1;
    std::ptr::copy_nonoverlapping(
        path_list.as_ptr(),
        drop_ptr.add(dropfiles_size) as *mut u16,
        path_list.len(),
    );
    GlobalUnlock(hdrop);

    let effect_format = preferred_drop_effect_format();
    let heffect = GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, std::mem::size_of::<u32>());
    if heffect.is_null() {
        GlobalFree(hdrop);
        return Err(io::Error::last_os_error());
    }
    let effect_ptr = GlobalLock(heffect) as *mut u32;
    if effect_ptr.is_null() {
        GlobalFree(heffect);
        GlobalFree(hdrop);
        return Err(io::Error::last_os_error());
    }
    *effect_ptr = match operation {
        ClipboardOperation::Copy => DROPEFFECT_COPY_VALUE,
        ClipboardOperation::Move => DROPEFFECT_MOVE_VALUE,
    };
    GlobalUnlock(heffect);

    if OpenClipboard(owner) == 0 {
        GlobalFree(heffect);
        GlobalFree(hdrop);
        return Err(io::Error::last_os_error());
    }

    let mut result = Ok(());
    let mut hdrop_owned = true;
    let mut heffect_owned = true;
    if EmptyClipboard() == 0 {
        result = Err(io::Error::last_os_error());
    } else if SetClipboardData(CF_HDROP_FORMAT, hdrop).is_null() {
        result = Err(io::Error::last_os_error());
    } else {
        hdrop_owned = false;
        if SetClipboardData(effect_format, heffect).is_null() {
            result = Err(io::Error::last_os_error());
        } else {
            heffect_owned = false;
        }
    }
    CloseClipboard();

    if result.is_err() {
        if heffect_owned {
            GlobalFree(heffect);
        }
        if hdrop_owned {
            GlobalFree(hdrop);
        }
        return result.map(|_| 0);
    }

    Ok(GetClipboardSequenceNumber())
}

unsafe fn read_file_paths_from_clipboard(
    owner: HWND,
) -> io::Result<Option<(Vec<PathBuf>, ClipboardOperation)>> {
    if OpenClipboard(owner) == 0 {
        return Err(io::Error::last_os_error());
    }

    let mut result = None;
    if IsClipboardFormatAvailable(CF_HDROP_FORMAT) != 0 {
        let handle = GetClipboardData(CF_HDROP_FORMAT);
        if !handle.is_null() {
            let count = DragQueryFileW(handle as _, 0xFFFF_FFFF, null_mut(), 0);
            let mut paths = Vec::with_capacity(count as usize);
            for index in 0..count {
                let length = DragQueryFileW(handle as _, index, null_mut(), 0);
                if length == 0 {
                    continue;
                }
                let mut buffer = vec![0u16; length as usize + 1];
                let copied =
                    DragQueryFileW(handle as _, index, buffer.as_mut_ptr(), buffer.len() as u32);
                let text = String::from_utf16_lossy(&buffer[..copied as usize]);
                if !text.is_empty() {
                    paths.push(PathBuf::from(text));
                }
            }
            let effect_format = preferred_drop_effect_format();
            let operation = if effect_format != 0 && IsClipboardFormatAvailable(effect_format) != 0
            {
                let effect_handle = GetClipboardData(effect_format);
                if !effect_handle.is_null() {
                    let effect_ptr = GlobalLock(effect_handle) as *const u32;
                    let effect = if !effect_ptr.is_null() {
                        let value = *effect_ptr;
                        GlobalUnlock(effect_handle);
                        value
                    } else {
                        DROPEFFECT_COPY_VALUE
                    };
                    if effect & DROPEFFECT_MOVE_VALUE != 0 {
                        ClipboardOperation::Move
                    } else {
                        ClipboardOperation::Copy
                    }
                } else {
                    ClipboardOperation::Copy
                }
            } else {
                ClipboardOperation::Copy
            };
            result = Some((paths, operation));
        }
    }
    CloseClipboard();
    Ok(result)
}

fn loword(value: u32) -> u16 {
    (value & 0xffff) as u16
}

fn hiword(value: u32) -> u16 {
    ((value >> 16) & 0xffff) as u16
}

fn ctrl_pressed() -> bool {
    unsafe { GetKeyState(VK_CONTROL as i32) < 0 }
}

fn shift_pressed() -> bool {
    unsafe { GetKeyState(VK_SHIFT as i32) < 0 }
}

unsafe fn read_window_text(hwnd: HWND) -> String {
    let len = GetWindowTextLengthW(hwnd);
    if len <= 0 {
        return String::new();
    }

    let mut buffer = vec![0u16; len as usize + 1];
    let copied = GetWindowTextW(hwnd, buffer.as_mut_ptr(), buffer.len() as i32);
    String::from_utf16_lossy(&buffer[..copied as usize])
}

fn widestr_ptr_to_string(ptr: *const u16) -> String {
    if ptr.is_null() {
        return String::new();
    }
    let mut len = 0usize;
    unsafe {
        while *ptr.add(len) != 0 {
            len += 1;
        }
        String::from_utf16_lossy(std::slice::from_raw_parts(ptr, len))
    }
}

fn display_name(path: &Path) -> String {
    path.file_name()
        .map(|name| name.to_string_lossy().to_string())
        .unwrap_or_else(|| path.display().to_string())
}

fn summarize_targets(targets: &[PathBuf]) -> String {
    match targets {
        [] => "brak elementów".to_string(),
        [single] => display_name(single),
        many => format!(
            "{}, pierwszy: {}",
            pluralized_elements(many.len()),
            display_name(&many[0])
        ),
    }
}

fn last_error_message(prefix: &str) -> String {
    format!("{prefix}: {}", io::Error::last_os_error())
}
