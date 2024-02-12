import ctypes, string
import shutil
import sys
import os

import win32api
from PyQt5.QtCore import QUrl, Qt, pyqtSignal, QMutex, QObject, QSize, QRunnable, \
    QThreadPool, QTimer
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QToolBar, QPushButton, \
    QHBoxLayout, QProgressBar, QListWidgetItem, QLabel, QListWidget, QGridLayout
from PyQt5.QtWebEngineWidgets import QWebEngineView, QWebEngineProfile
from PyQt5.QtGui import QIcon, QFont, QPixmap
from pytube import YouTube

class DownloaderSignals(QObject):
    progress = pyqtSignal(int,int)
    name = pyqtSignal(int,str)
    finished = pyqtSignal(str)

class Downloader(QRunnable):

    def __init__(self, url, index, output_folder):
        super().__init__()
        self.url = url
        self.index = index
        self.output_folder = output_folder
        self.signals = DownloaderSignals()
        self.progress = 0

    def run(self):
        try:
            video = YouTube(self.url, on_progress_callback=self.update_progress)
            audio_stream = video.streams.filter(only_audio=True).first()
            song_name = audio_stream.title
            song_name = song_name.replace('/', '').replace('\\', '').replace('.', '')
            file_name = os.path.join(self.output_folder, song_name + '.mp3')
            if not os.path.exists(file_name):
                self.signals.name.emit(self.index, song_name)
                audio_stream.download(filename=file_name)
            else:
                song_name = ''
        except Exception as e:
            print("Error:", str(e))
        finally:
            if song_name and self.progress == 0:
                self.signals.progress.emit(self.index, 100)
            self.signals.finished.emit(song_name)

    def update_progress(self, stream, chunk, remaining):
        total_size = stream.filesize
        bytes_downloaded = total_size - remaining
        self.progress = int(bytes_downloaded / total_size * 100)
        self.signals.progress.emit(self.index, self.progress)

class YouTubeViewer(QMainWindow):
    def __init__(self):
        super().__init__()

        self.mutex = QMutex()
        self.progress_bars = []
        self.name_labels = []
        self.setWindowTitle("YouTube Downloader")
        self.base_folder = os.path.dirname(os.path.abspath(__file__))
        self.images_folder = os.path.join(self.base_folder, 'src', 'images')
        app_icon = QIcon(os.path.join(self.images_folder, 'app_icon.png'))
        self.setWindowIcon(app_icon)
        screen_geometry = QApplication.desktop().availableGeometry()
        initial_width = int(screen_geometry.width())
        initial_height = int(screen_geometry.height())
        initial_x = (screen_geometry.width() - initial_width) // 2
        initial_y = (screen_geometry.height() - initial_height) // 2

        self.setGeometry(0, 0, initial_width, initial_height)

        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        # Use QHBoxLayout for the main layout
        layout = QHBoxLayout(central_widget)
        layout.setSpacing(0)

        # Add the webview to the left side
        webview = QWebEngineView()
        layout.addWidget(webview)

        toolbar = QToolBar(self)
        self.addToolBar(toolbar)

        # Set icons for actions
        back_icon = QIcon(os.path.join(self.images_folder, 'back_icon.png'))
        download_icon = QIcon(os.path.join(self.images_folder, 'download_icon.png'))
        transfer_icon = QIcon(os.path.join(self.images_folder, 'transfer_icon.png'))
        self.pendrive_green_icon = QIcon(os.path.join(self.images_folder, 'pendrive_green.png'))
        self.pendrive_grey_icon = QIcon(os.path.join(self.images_folder, 'pendrive_grey.png'))

        back_action = QPushButton(back_icon, "返回", self)
        back_action.clicked.connect(self.go_back)
        toolbar.addWidget(back_action)

        download_button = QPushButton(download_icon, "下载", self)
        download_button.clicked.connect(self.download_mp3)
        toolbar.addWidget(download_button)

        transfer_button = QPushButton(transfer_icon, "转移", self)
        transfer_button.clicked.connect(self.transfer_song)
        toolbar.addWidget(transfer_button)
        toolbar.setIconSize(QSize(256, 256))

        # Apply style sheets for hover and clicked effects
        self.apply_style_sheet(back_action)
        self.apply_style_sheet(download_button)
        self.apply_style_sheet(transfer_button)

        # Add an instance of SidePanel to the side panel
        self.side_panel = QWidget(self)
        self.side_panel.setFixedWidth(int(self.width() * 0.15))
        layout.addWidget(self.side_panel)

        # Create a layout for the side panel
        self.side_layout = QVBoxLayout(self.side_panel)
        self.side_layout.setContentsMargins(int(self.side_panel.width()*0.05),0,0,0)
        # Create a widget for displaying download information

        self.side_layout.addWidget(QLabel("下载列表"))
        self.download_info_widget = QListWidget(self)
        self.download_info_widget.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.side_layout.addWidget(self.download_info_widget)

        self.pendrive = QPushButton(self.pendrive_grey_icon, '', self)
        self.pendrive.setIconSize(QSize(56,56))
        self.pendrive.font().setPointSize(12)
        self.pendrive.setStyleSheet("text-align: left;padding-left: 10px")
        self.pendrive.clicked.connect(self.load_pendrive)
        self.side_layout.addWidget(self.pendrive)
        self.load_pendrive()

        # Create a widget for listing all MP3 files in a folder

        self.side_layout.addWidget(QLabel("已下载的歌曲"))
        self.mp3_list_widget = QListWidget(self)
        self.mp3_list_widget.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.side_layout.addWidget(self.mp3_list_widget)
        self.output_folder = os.path.join(self.base_folder, "output_mp3")
        self.update_mp3_list()

        delete_button = QPushButton('删除')
        delete_button.clicked.connect(self.delete_mp3_list)
        self.side_layout.addWidget(delete_button)

        self.threadpool = QThreadPool()

        QWebEngineProfile.defaultProfile().setHttpAcceptLanguage('zh-CN')
        webview.setUrl(QUrl("https://www.youtube.com"))
        self.webview = webview

    def go_back(self):
        if self.webview.history().canGoBack():
            self.webview.back()

    def download_mp3(self):
        current_url = self.webview.url().toString()
        current_url = 'https://www.youtube.com/embed/Rs6j23OMwMs'
        if current_url and current_url != 'https://www.youtube.com':
            index = len(self.progress_bars)
            download_thread = Downloader(current_url, index, self.output_folder)
            download_thread.signals.progress.connect(self.update_progress_bar)
            download_thread.signals.name.connect(self.update_name_label)
            download_thread.signals.finished.connect(self.add_mp3_list)
            self.threadpool.start(download_thread)

    def update_progress_bar(self, index, progress):
        self.mutex.lock()
        self.progress_bars[index].setValue(progress)
        self.mutex.unlock()

    def update_name_label(self, index, name):
        self.mutex.lock()
        progress_bar, name_label = self.add_download_info()
        self.progress_bars.append(progress_bar)
        self.name_labels.append(name_label)
        self.name_labels[index].setText(name)
        self.mutex.unlock()

    def transfer_song(self):
        if self.pendrive_path:
            output_mp3 = os.path.join(self.base_folder, 'output_mp3')
            mp3_files = [mp3 for mp3 in os.listdir(output_mp3) if mp3.endswith('.mp3')]
            for mp3 in mp3_files:
                source_path = os.path.join(output_mp3, mp3)
                try:
                    # Check if the file exists
                    if not os.path.exists(source_path):
                        print(f"The file '{source_path}' does not exist.")
                        return
                    # Check if the pendrive path is valid
                    if not os.path.exists(self.pendrive_path):
                        print(f"The pendrive path '{self.pendrive_path}' does not exist.")
                        return
                    # Move the file to the pendrive
                    shutil.move(source_path, self.pendrive_path)
                    print(f"File '{os.path.basename(source_path)}' successfully moved to pendrive.")
                except Exception as e:
                    print(f"An error occurred: {e}")
            self.mp3_list_widget.clear()

    def load_pendrive(self):
        bitmask = ctypes.windll.kernel32.GetLogicalDrives()
        drives = [letter for i, letter in enumerate(string.ascii_uppercase) if bitmask & (1 << i)]
        ext_drives = [letter for letter in drives if ctypes.windll.kernel32.GetDriveTypeW(letter + ':') == 2]  # DRIVE_REMOVABLE = 2
        if len(ext_drives) > 0:
            self.pendrive_path = ext_drives[0] + ':\\'
            drive_info = win32api.GetVolumeInformation(self.pendrive_path)
            self.pendrive.setText(self.pendrive_path + drive_info[0])
            self.pendrive.setIcon(self.pendrive_green_icon)
        else:
            self.pendrive.setText('')
            self.pendrive.setIcon(self.pendrive_grey_icon)

    def add_download_info(self):
        icon_path = os.path.join(self.images_folder, 'music.png')
        item = QListWidgetItem(self.download_info_widget)
        item.setSizeHint(QSize(self.download_info_widget.width(), int(self.download_info_widget.height()*0.2)))
        download_widget = QWidget()

        # Create labels for name, icon, and progress
        name_label = QLabel()
        name_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)  # Align text to the left

        icon_label = QLabel()
        icon_label.setPixmap(QPixmap(icon_path).scaled(10, 10, Qt.KeepAspectRatio))

        progress_bar = QProgressBar()

        # Add labels to the download widget layout
        download_layout = QGridLayout(download_widget)
        download_layout.addWidget(icon_label, 0, 0, 1, 1)
        download_layout.addWidget(name_label, 0, 1, 1, 3)
        download_layout.addWidget(progress_bar, 1, 0, 1, 10)

        self.download_info_widget.addItem(item)
        self.download_info_widget.setItemWidget(item, download_widget)

        return progress_bar, name_label

    def update_mp3_list(self):
        if not os.path.exists(self.output_folder):
            os.makedirs(self.output_folder)
        else:
            mp3_files = os.listdir(self.output_folder)
            for mp3_file in mp3_files:
                item = QListWidgetItem((mp3_file.split("/"))[-1].split('.')[0])
                self.mp3_list_widget.addItem(item)

    def add_mp3_list(self, name):
        if not name: return
        self.mutex.lock()
        item = QListWidgetItem(name)
        self.mp3_list_widget.addItem(item)
        self.mutex.unlock()

    def delete_mp3_list(self):
        selected_mp3 = self.mp3_list_widget.selectedItems()
        if not selected_mp3: return
        for mp3 in selected_mp3:
            index = self.mp3_list_widget.row(mp3)
            mp3_name = mp3.text()
            self.mp3_list_widget.takeItem(index)
            try:
                os.remove(os.path.join(self.output_folder, mp3_name + '.mp3'))
            except Exception as e:
                print("Error:", str(e))

    def apply_style_sheet(self, widget):
        widget.setMinimumSize(int(0.08 * self.width()), int(0.02 * self.height()))

        # Adjust text font size
        widget.setFont(QFont("Microsoft YaHei", int(0.007 * self.width())))

        # Adjust icon size
        widget.setIconSize(QSize(int(0.03 * self.width()), int(0.03 * self.width())))

        # Apply modern styling
        widget.setStyleSheet(
            """
            QPushButton {
                border: none;
                border-radius: 5px;
                padding: 5px;
            }

            QPushButton:hover {
                background-color: lightgray;
            }

            QPushButton:pressed {
                background-color: gray;
            }
            """
        )


if __name__ == "__main__":
    app = QApplication(sys.argv)
    viewer = YouTubeViewer()
    viewer.show()
    sys.exit(app.exec_())
