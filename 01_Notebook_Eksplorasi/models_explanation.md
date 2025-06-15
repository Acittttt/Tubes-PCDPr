# Penjelasan Model pada Proyek MediaPipe Head Gesture Control

## Apakah Ada File Model (.h5 / .pt / .pb / .rknn) di Eksplorasi Ini?

**Tidak ada file model eksternal** yang digunakan/di-attach pada eksplorasi ini. Semua proses deteksi landmark wajah, pose tubuh, dan holistic pada notebook menggunakan **model bawaan (built-in) dari library MediaPipe**. Model ini sudah pre-trained oleh Google dan langsung diakses dari package MediaPipe, sehingga:

- Tidak ada file model (.h5/.pt/.pb/.rknn) yang perlu diupload, di-train, atau didokumentasikan terpisah.
- Semua pemanggilan model terjadi secara otomatis saat memanggil `mp.solutions.face_mesh`, `mp.solutions.pose`, atau `mp.solutions.holistic` di dalam kode Python.

---

## Model yang Digunakan (Sesuai Kode di Notebook)

### 1. MediaPipe Face Mesh

- **Deskripsi:** Model untuk mendeteksi 468 landmark wajah secara real-time.
- **Tipe Model:** Built-in, pre-trained, tidak ada file eksternal.
- **Contoh Pemanggilan di Notebook:**
  ```python
  mp_face_mesh = mp.solutions.face_mesh
  with mp_face_mesh.FaceMesh(
      max_num_faces=1,
      refine_landmarks=True,
      min_detection_confidence=0.5,
      min_tracking_confidence=0.5
  ) as face_mesh:
      ...
  ```

### 2. MediaPipe Pose

- **Deskripsi:** Model untuk mendeteksi 33 landmark tubuh (full body/pose).
- **Tipe Model:** Built-in, pre-trained, tidak ada file eksternal.
- **Contoh Pemanggilan di Notebook:**
  ```python
  mp_pose = mp.solutions.pose
  with mp_pose.Pose(
      min_detection_confidence=0.5,
      min_tracking_confidence=0.5,
      model_complexity=1
  ) as pose:
      ...
  ```

### 3. MediaPipe Holistic

- **Deskripsi:** Model integrasi Face Mesh + Pose + Hands (543 landmark total).
- **Tipe Model:** Built-in, pre-trained, tidak ada file eksternal.
- **Contoh Pemanggilan di Notebook:**
  ```python
  mp_holistic = mp.solutions.holistic
  with mp_holistic.Holistic(
      min_detection_confidence=0.5,
      min_tracking_confidence=0.5,
      model_complexity=1,
      refine_face_landmarks=True
  ) as holistic:
      ...
  ```

---

## Penjelasan Kenapa Tidak Ada File Model

- Semua model di atas **sudah di-bundle dalam library MediaPipe** (diunduh otomatis saat install).
- Tidak memerlukan file `.h5`, `.pt`, `.pb`, atau `.rknn` terpisah.
- Model digunakan secara “black box” oleh MediaPipe, tidak bisa di-train ulang/customize dari sisi pengguna.

---

## Referensi Resmi

- [MediaPipe Face Mesh Documentation](https://ai.google.dev/edge/mediapipe/solutions/vision/face_landmarker)
- [MediaPipe Pose Documentation](https://ai.google.dev/edge/mediapipe/solutions/vision/pose_landmarker)
- [MediaPipe Holistic Documentation](https://ai.google.dev/edge/mediapipe/solutions/vision/holistic)

---

**Kesimpulan:**  
Struktur folder dan file notebook sudah sesuai requirement tugas, walaupun tidak ada file model terpisah, karena MediaPipe tidak memerlukan itu dan notebook tetap dapat dijalankan serta dieksplorasi sepenuhnya.
