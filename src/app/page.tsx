'use client';

import React, { useState, useEffect, useRef } from 'react';
import {
  Home,
  BookOpen,
  Users,
  Settings,
  Plus,
  Search,
  Scan,
  Trash2,
  Edit,
  Download,
  Upload,
  CheckCircle2,
  Book,
  ChevronRight,
  Star,
  X,
  FileSpreadsheet,
  FileText,
  Calendar,
  GripVertical,
  Layers,
  Sparkles,
  Camera,
  RotateCcw,
  Zap,
  ZapOff,
  AlertCircle,
  Loader2,
} from 'lucide-react';

import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import { Html5Qrcode } from 'html5-qrcode';

// ============================================================
// INTERFACCIA DATI LIBRO
// ============================================================

export interface BookItem {
  id: string;
  title: string;
  author: string;
  publishCountry?: string;
  coverUrl?: string;
  publisher?: string;
  publishYear?: string;
  pages?: number;
  genre?: string;
  seriesTag?: string;
  volume?: string;
  isClassic?: boolean;
  format: 'cartaceo' | 'ebook' | 'ebook_and_paper';
  isRead: boolean;
  readMonth?: string;
  readYear?: number;
  readMonthYear?: string;
  rating?: number;
  notes?: string;
  isbn?: string;
  createdAt: number;
}

const MONTHS = [
  'Gennaio',
  'Febbraio',
  'Marzo',
  'Aprile',
  'Maggio',
  'Giugno',
  'Luglio',
  'Agosto',
  'Settembre',
  'Ottobre',
  'Novembre',
  'Dicembre',
];

// ============================================================
// COMPONENTE SCANNER ISBN
// ============================================================

interface ISBNScannerProps {
  onDetected: (isbn: string) => void;
  onClose: () => void;
}

function ISBNScanner({ onDetected, onClose }: ISBNScannerProps) {
  const scannerRef = useRef<Html5Qrcode | null>(null);
  const startingRef = useRef(false);
  const detectedRef = useRef(false);

  const [scannerStarted, setScannerStarted] = useState(false);
  const [scannerError, setScannerError] = useState('');
  const [isLoading, setIsLoading] = useState(true);
  const [cameraFacing, setCameraFacing] = useState<'environment' | 'user'>(
    'environment'
  );

  const [torchSupported, setTorchSupported] = useState(false);
  const [torchOn, setTorchOn] = useState(false);

  const scannerElementId = 'isbn-reader';

  // ----------------------------------------------------------
  // Ferma scanner
  // ----------------------------------------------------------

  const stopScanner = async () => {
    const scanner = scannerRef.current;

    if (!scanner) return;

    try {
      const state = scanner.getState();

      // 2 = SCANNING
      // 3 = PAUSED
      if (state === 2 || state === 3) {
        await scanner.stop();
      }
    } catch (error) {
      console.warn('Errore arrestando lo scanner:', error);
    }

    try {
      scanner.clear();
    } catch (error) {
      console.warn('Errore clear scanner:', error);
    }

    scannerRef.current = null;
    setScannerStarted(false);
  };

  // ----------------------------------------------------------
  // Normalizza ISBN
  // ----------------------------------------------------------

  const normalizeISBN = (value: string): string => {
    return value.replace(/[^0-9Xx]/g, '').toUpperCase();
  };

  // ----------------------------------------------------------
  // Verifica ISBN-13
  // ----------------------------------------------------------

  const isValidISBN13 = (isbn: string): boolean => {
    const clean = normalizeISBN(isbn);

    if (!/^\d{13}$/.test(clean)) {
      return false;
    }

    let sum = 0;

    for (let i = 0; i < 12; i++) {
      const digit = Number(clean[i]);
      sum += i % 2 === 0 ? digit : digit * 3;
    }

    const checkDigit = (10 - (sum % 10)) % 10;

    return checkDigit === Number(clean[12]);
  };

  // ----------------------------------------------------------
  // Controlla supporto torcia
  // ----------------------------------------------------------

  const checkTorchSupport = () => {
    try {
      const video = document.querySelector(
        '#isbn-reader video'
      ) as HTMLVideoElement | null;

      const stream = video?.srcObject as MediaStream | null;

      const track = stream?.getVideoTracks?.()[0];

      if (!track) {
        setTorchSupported(false);
        return;
      }

      const capabilities = track.getCapabilities?.();

      if (capabilities && 'torch' in capabilities) {
        setTorchSupported(true);
      } else {
        setTorchSupported(false);
      }
    } catch {
      setTorchSupported(false);
    }
  };

  // ----------------------------------------------------------
  // Torcia
  // ----------------------------------------------------------

  const toggleTorch = async () => {
    try {
      const video = document.querySelector(
        '#isbn-reader video'
      ) as HTMLVideoElement | null;

      const stream = video?.srcObject as MediaStream | null;

      const track = stream?.getVideoTracks?.()[0];

      if (!track) return;

      await track.applyConstraints({
        advanced: [{ torch: !torchOn } as MediaTrackConstraintSet],
      });

      setTorchOn(!torchOn);
    } catch (error) {
      console.warn('Torcia non disponibile:', error);
      setTorchSupported(false);
    }
  };

  // ----------------------------------------------------------
  // Avvia scanner
  // ----------------------------------------------------------

  const startScanner = async () => {
    if (startingRef.current) return;

    startingRef.current = true;
    detectedRef.current = false;

    setScannerError('');
    setIsLoading(true);

    try {
      if (typeof window === 'undefined') {
        return;
      }

      // ------------------------------------------------------
      // Controllo HTTPS
      // ------------------------------------------------------

      const isLocalhost =
        window.location.hostname === 'localhost' ||
        window.location.hostname === '127.0.0.1';

      if (!window.isSecureContext && !isLocalhost) {
        setScannerError(
          'La fotocamera richiede HTTPS. Apri questa applicazione tramite un indirizzo https://.'
        );

        setIsLoading(false);
        return;
      }

      // ------------------------------------------------------
      // Controllo getUserMedia
      // ------------------------------------------------------

      if (
        !navigator.mediaDevices ||
        !navigator.mediaDevices.getUserMedia
      ) {
        setScannerError(
          'Safari non permette di accedere alla fotocamera da questa pagina. Controlla che il sito sia aperto tramite HTTPS.'
        );

        setIsLoading(false);
        return;
      }

      // ------------------------------------------------------
      // Piccolo delay: necessario affinché il DOM del modal
      // sia disponibile
      // ------------------------------------------------------

      await new Promise((resolve) => setTimeout(resolve, 150));

      // ------------------------------------------------------
      // Istanza scanner
      // ------------------------------------------------------

      const scanner = new Html5Qrcode(scannerElementId);

      scannerRef.current = scanner;

      // ------------------------------------------------------
      // Configurazione
      // ------------------------------------------------------

      const config = {
        fps: 10,

        qrbox: (viewfinderWidth: number, viewfinderHeight: number) => {
          const width = Math.min(
            Math.floor(viewfinderWidth * 0.85),
            420
          );

          const height = Math.min(
            Math.floor(viewfinderHeight * 0.35),
            180
          );

          return {
            width,
            height,
          };
        },

        aspectRatio: 1.7777778,

        formatsToSupport: [
          // EAN-13 = ISBN-13
          0,
        ],
      };

      // ------------------------------------------------------
      // Avvio
      // ------------------------------------------------------

      await scanner.start(
        {
          facingMode: cameraFacing,
        },
        config,

        async (decodedText) => {
          if (detectedRef.current) {
            return;
          }

          const cleanCode = normalizeISBN(decodedText);

          // --------------------------------------------------
          // Lo scanner può leggere anche altri codici.
          // Accettiamo solo ISBN-13.
          // --------------------------------------------------

          if (!/^97\d{11}$/.test(cleanCode)) {
            return;
          }

          if (!isValidISBN13(cleanCode)) {
            return;
          }

          detectedRef.current = true;

          if (navigator.vibrate) {
            navigator.vibrate(100);
          }

          await stopScanner();

          onDetected(cleanCode);
        },

        () => {
          // Non facciamo nulla sugli errori di singolo frame.
          // html5-qrcode genera moltissimi errori durante la
          // normale ricerca del codice.
        }
      );

      setScannerStarted(true);
      setIsLoading(false);

      // ------------------------------------------------------
      // Aspettiamo che il video venga creato
      // ------------------------------------------------------

      setTimeout(() => {
        checkTorchSupport();
      }, 800);
    } catch (error: any) {
      console.error('Errore avvio scanner:', error);

      let message =
        'Impossibile accedere alla fotocamera.';

      if (error?.name === 'NotAllowedError') {
        message =
          'Accesso alla fotocamera negato. Su iPhone vai in Impostazioni > Safari > Fotocamera e consenti l’accesso.';
      } else if (error?.name === 'NotFoundError') {
        message =
          'Nessuna fotocamera disponibile sul dispositivo.';
      } else if (error?.name === 'NotReadableError') {
        message =
          'La fotocamera è già utilizzata da un’altra applicazione o scheda.';
      } else if (error?.name === 'OverconstrainedError') {
        message =
          'La fotocamera richiesta non è disponibile. Prova a riaprire lo scanner.';
      } else if (
        String(error?.message || '')
          .toLowerCase()
          .includes('permission')
      ) {
        message =
          'Safari non ha concesso l’accesso alla fotocamera.';
      }

      setScannerError(message);
      setIsLoading(false);

      await stopScanner();
    } finally {
      startingRef.current = false;
    }
  };

  // ----------------------------------------------------------
  // Cambio fotocamera
  // ----------------------------------------------------------

  const switchCamera = async () => {
    const newFacing =
      cameraFacing === 'environment' ? 'user' : 'environment';

    await stopScanner();

    setTorchOn(false);
    setCameraFacing(newFacing);

    setTimeout(() => {
      startScanner();
    }, 250);
  };

  // ----------------------------------------------------------
  // Mount
  // ----------------------------------------------------------

  useEffect(() => {
    startScanner();

    return () => {
      stopScanner();
    };

    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // ----------------------------------------------------------
  // Chiusura scanner
  // ----------------------------------------------------------

  const handleClose = async () => {
    await stopScanner();
    onClose();
  };

  return (
    <div className="fixed inset-0 z-[70] bg-black flex flex-col">
      {/* HEADER */}

      <div className="absolute top-0 left-0 right-0 z-20 bg-gradient-to-b from-black/80 to-transparent px-5 pt-12 pb-8">
        <div className="flex items-center justify-between">
          <div>
            <h2 className="text-lg font-serif font-bold text-white">
              Scansiona ISBN
            </h2>

            <p className="text-[11px] text-white/70 mt-1">
              Inquadra il codice a barre sul retro del libro
            </p>
          </div>

          <button
            type="button"
            onClick={handleClose}
            className="w-10 h-10 rounded-full bg-white/15 backdrop-blur-md flex items-center justify-center text-white"
          >
            <X className="w-5 h-5" />
          </button>
        </div>
      </div>

      {/* CAMERA */}

      <div className="flex-1 flex items-center justify-center bg-black relative overflow-hidden">
        <div
          id={scannerElementId}
          className="w-full h-full"
        />

        {/* Overlay centrale */}

        {!scannerError && (
          <div className="absolute inset-0 pointer-events-none flex items-center justify-center">
            <div className="relative w-[88%] max-w-[430px] h-[190px]">
              {/* angoli */}

              <div className="absolute top-0 left-0 w-10 h-10 border-t-4 border-l-4 border-amber-400 rounded-tl-2xl" />

              <div className="absolute top-0 right-0 w-10 h-10 border-t-4 border-r-4 border-amber-400 rounded-tr-2xl" />

              <div className="absolute bottom-0 left-0 w-10 h-10 border-b-4 border-l-4 border-amber-400 rounded-bl-2xl" />

              <div className="absolute bottom-0 right-0 w-10 h-10 border-b-4 border-r-4 border-amber-400 rounded-br-2xl" />

              {/* linea scanner */}

              {scannerStarted && (
                <div className="absolute left-5 right-5 top-1/2 h-0.5 bg-amber-400 shadow-[0_0_12px_rgba(251,191,36,0.9)] animate-pulse" />
              )}
            </div>
          </div>
        )}

        {/* Loading */}

        {isLoading && !scannerError && (
          <div className="absolute inset-0 bg-black/30 flex items-center justify-center">
            <div className="bg-black/70 backdrop-blur-md rounded-2xl px-5 py-4 flex items-center gap-3">
              <Loader2 className="w-5 h-5 text-amber-400 animate-spin" />

              <span className="text-white text-xs font-semibold">
                Avvio fotocamera...
              </span>
            </div>
          </div>
        )}

        {/* Errore */}

        {scannerError && (
          <div className="absolute inset-0 flex items-center justify-center p-6">
            <div className="bg-[#FFFDF9] rounded-3xl p-6 max-w-sm w-full shadow-2xl">
              <div className="w-12 h-12 rounded-2xl bg-rose-100 text-rose-700 flex items-center justify-center mx-auto mb-4">
                <AlertCircle className="w-6 h-6" />
              </div>

              <h3 className="text-base font-serif font-bold text-amber-950 text-center">
                Fotocamera non disponibile
              </h3>

              <p className="text-xs text-amber-900/70 text-center mt-2 leading-relaxed">
                {scannerError}
              </p>

              <button
                type="button"
                onClick={() => {
                  setScannerError('');
                  startScanner();
                }}
                className="w-full mt-5 py-3 bg-amber-800 text-white rounded-xl text-xs font-bold flex items-center justify-center gap-2"
              >
                <RotateCcw className="w-4 h-4" />
                Riprova
              </button>

              <button
                type="button"
                onClick={handleClose}
                className="w-full mt-2 py-3 bg-amber-100 text-amber-950 rounded-xl text-xs font-bold"
              >
                Inserisci ISBN manualmente
              </button>
            </div>
          </div>
        )}
      </div>

      {/* CONTROLLI */}

      <div className="absolute bottom-0 left-0 right-0 z-20 bg-gradient-to-t from-black/90 via-black/70 to-transparent px-5 pt-10 pb-10">
        <p className="text-center text-white/80 text-xs font-medium mb-5">
          Cerca un codice ISBN-13 che inizi con 978 o 979
        </p>

        <div className="flex items-center justify-center gap-4">
          {/* Torcia */}

          {torchSupported && (
            <button
              type="button"
              onClick={toggleTorch}
              className="w-12 h-12 rounded-full bg-white/15 backdrop-blur-md flex items-center justify-center text-white"
            >
              {torchOn ? (
                <ZapOff className="w-5 h-5" />
              ) : (
                <Zap className="w-5 h-5" />
              )}
            </button>
          )}

          {/* Switch camera */}

          <button
            type="button"
            onClick={switchCamera}
            className="w-14 h-14 rounded-full bg-white text-amber-950 flex items-center justify-center shadow-xl"
          >
            <RotateCcw className="w-6 h-6" />
          </button>

          {/* Chiudi */}

          <button
            type="button"
            onClick={handleClose}
            className="w-12 h-12 rounded-full bg-white/15 backdrop-blur-md flex items-center justify-center text-white"
          >
            <X className="w-5 h-5" />
          </button>
        </div>
      </div>
    </div>
  );
}

// ============================================================
// APP PRINCIPALE
// ============================================================

export default function LibraryApp() {
  const [activeTab, setActiveTab] = useState<
    'home' | 'read' | 'authors' | 'settings'
  >('home');

  const [books, setBooks] = useState<BookItem[]>([]);
  const [isLoaded, setIsLoaded] = useState(false);

  // ----------------------------------------------------------
  // Modali
  // ----------------------------------------------------------

  const [isAddModalOpen, setIsAddModalOpen] = useState(false);

  const [isScannerOpen, setIsScannerOpen] = useState(false);

  const [selectedBookDetail, setSelectedBookDetail] =
    useState<BookItem | null>(null);

  const [selectedAuthor, setSelectedAuthor] =
    useState<string | null>(null);

  const [homeSubView, setHomeSubView] = useState<
    'none' | 'classics' | 'genres'
  >('none');

  const [selectedGenreHome, setSelectedGenreHome] =
    useState<string | null>(null);

  // ----------------------------------------------------------
  // Drag
  // ----------------------------------------------------------

  const [draggedIndex, setDraggedIndex] =
    useState<number | null>(null);

  // ----------------------------------------------------------
  // ISBN
  // ----------------------------------------------------------

  const [isbnInput, setIsbnInput] = useState('');

  const [isSearchingIsbn, setIsSearchingIsbn] =
    useState(false);

  // ----------------------------------------------------------
  // Filtri
  // ----------------------------------------------------------

  const [filterGenre, setFilterGenre] =
    useState<string>('all');

  const [filterFormat, setFilterFormat] =
    useState<string>('all');

  const [filterYear, setFilterYear] =
    useState<string>('all');

  const [searchQuery, setSearchQuery] =
    useState('');

  // ----------------------------------------------------------
  // Anni
  // ----------------------------------------------------------

  const currentYearNum = new Date().getFullYear();

  const startYear = 2023;

  const yearsList = Array.from(
    {
      length: Math.max(
        1,
        currentYearNum - startYear + 1
      ),
    },
    (_, i) => currentYearNum - i
  );

  // ----------------------------------------------------------
  // Form
  // ----------------------------------------------------------

  const [formData, setFormData] =
    useState<Partial<BookItem>>({
      title: '',
      author: '',
      publishCountry: '',
      coverUrl: '',
      publisher: '',
      publishYear: '',
      pages: undefined,
      genre: '',
      seriesTag: '',
      volume: '',
      isClassic: false,
      format: 'cartaceo',
      isRead: false,
      readMonth: 'Gennaio',
      readYear: currentYearNum,
      rating: 5,
      notes: '',
      isbn: '',
    });

  // ==========================================================
  // CARICAMENTO
  // ==========================================================

  useEffect(() => {
    const saved = localStorage.getItem(
      'ios_library_books_v5'
    );

    if (saved) {
      try {
        setBooks(JSON.parse(saved));
      } catch (e) {
        console.error(
          'Errore caricamento storage',
          e
        );
      }
    }

    setIsLoaded(true);
  }, []);

  // ==========================================================
  // SALVATAGGIO
  // ==========================================================

  useEffect(() => {
    if (isLoaded) {
      localStorage.setItem(
        'ios_library_books_v5',
        JSON.stringify(books)
      );
    }
  }, [books, isLoaded]);

  // ==========================================================
  // STATISTICHE
  // ==========================================================

  const totalBooks = books.length;

  const totalCartacei = books.filter(
    (b) =>
      b.format === 'cartaceo' ||
      b.format === 'ebook_and_paper'
  ).length;

  const totalEbook = books.filter(
    (b) =>
      b.format === 'ebook' ||
      b.format === 'ebook_and_paper'
  ).length;

  const readBooks = books.filter(
    (b) => b.isRead
  );

  const readBooksCount = readBooks.length;

  const readCartacei = readBooks.filter(
    (b) =>
      b.format === 'cartaceo' ||
      b.format === 'ebook_and_paper'
  ).length;

  const readEbook = readBooks.filter(
    (b) =>
      b.format === 'ebook' ||
      b.format === 'ebook_and_paper'
  ).length;

  const readThisYearBooks = books.filter(
    (b) => {
      if (!b.isRead) return false;

      if (b.readYear) {
        return b.readYear === currentYearNum;
      }

      if (
        b.readMonthYear &&
        b.readMonthYear.includes(
          currentYearNum.toString()
        )
      ) {
        return true;
      }

      return false;
    }
  );

  const readThisYearCount =
    readThisYearBooks.length;

  const readThisYearCartacei =
    readThisYearBooks.filter(
      (b) =>
        b.format === 'cartaceo' ||
        b.format === 'ebook_and_paper'
    ).length;

  const readThisYearEbook =
    readThisYearBooks.filter(
      (b) =>
        b.format === 'ebook' ||
        b.format === 'ebook_and_paper'
    ).length;

  // ==========================================================
  // GOOGLE BOOKS
  // ==========================================================

  const handleSearchBookByISBN = async (
    codeToSearch?: string
  ) => {
    const query = (
      codeToSearch || isbnInput
    ).replace(/[^0-9Xx]/g, '');

    if (!query) {
      alert('Inserisci un codice ISBN.');
      return;
    }

    setIsSearchingIsbn(true);

    try {
      const res = await fetch(
        `https://www.googleapis.com/books/v1/volumes?q=isbn:${encodeURIComponent(
          query
        )}`
      );

      if (!res.ok) {
        throw new Error(
          `HTTP ${res.status}`
        );
      }

      const data = await res.json();

      if (
        data.items &&
        data.items.length > 0
      ) {
        const info =
          data.items[0].volumeInfo;

        const image =
          info.imageLinks?.thumbnail ||
          info.imageLinks?.smallThumbnail ||
          '';

        setFormData((prev) => ({
          ...prev,

          title:
            info.title ||
            prev.title ||
            '',

          author:
            info.authors?.join(', ') ||
            prev.author ||
            '',

          publisher:
            info.publisher ||
            prev.publisher ||
            '',

          publishYear:
            info.publishedDate
              ? info.publishedDate.substring(
                  0,
                  4
                )
              : prev.publishYear || '',

          pages:
            info.pageCount ||
            prev.pages ||
            undefined,

          genre:
            info.categories?.[0] ||
            prev.genre ||
            '',

          coverUrl:
            image
              ? image.replace(
                  'http:',
                  'https:'
                )
              : prev.coverUrl || '',

          isbn: query,
        }));

        setIsbnInput(query);
      } else {
        alert(
          'Nessun dato trovato per questo ISBN. Puoi compilare i dettagli manualmente.'
        );

        setIsbnInput(query);

        setFormData((prev) => ({
          ...prev,
          isbn: query,
        }));
      }
    } catch (err) {
      console.error(
        'Errore Google Books:',
        err
      );

      alert(
        'Errore durante la ricerca online. Controlla la connessione e riprova.'
      );
    } finally {
      setIsSearchingIsbn(false);
    }
  };

  // ==========================================================
  // ISBN SCANSIONATO
  // ==========================================================

  const handleISBNDetected = async (
    isbn: string
  ) => {
    setIsScannerOpen(false);

    setIsbnInput(isbn);

    setFormData((prev) => ({
      ...prev,
      isbn,
    }));

    // Aspettiamo un istante per chiudere
    // correttamente la fotocamera
    await new Promise((resolve) =>
      setTimeout(resolve, 150)
    );

    await handleSearchBookByISBN(isbn);
  };

  // ==========================================================
  // SALVA LIBRO
  // ==========================================================

  const handleSaveBook = (
    e: React.FormEvent
  ) => {
    e.preventDefault();

    if (
      !formData.title ||
      !formData.author
    ) {
      alert(
        'Inserisci almeno Titolo e Autore.'
      );

      return;
    }

    const formattedReadMonthYear =
      formData.isRead &&
      formData.readMonth &&
      formData.readYear
        ? `${formData.readMonth} ${formData.readYear}`
        : '';

    const newBook: BookItem = {
      id:
        formData.id ||
        Date.now().toString(),

      title: formData.title || '',

      author: formData.author || '',

      publishCountry:
        formData.publishCountry || '',

      coverUrl:
        formData.coverUrl || '',

      publisher:
        formData.publisher || '',

      publishYear:
        formData.publishYear || '',

      pages:
        Number(formData.pages) ||
        undefined,

      genre:
        formData.genre || '',

      seriesTag:
        formData.seriesTag || '',

      volume:
        formData.volume || '',

      isClassic:
        !!formData.isClassic,

      format:
        formData.format ||
        'cartaceo',

      isRead:
        !!formData.isRead,

      readMonth:
        formData.readMonth ||
        'Gennaio',

      readYear:
        formData.readYear ||
        currentYearNum,

      readMonthYear:
        formattedReadMonthYear,

      rating:
        Number(formData.rating) || 5,

      notes:
        formData.notes || '',

      isbn:
        formData.isbn || '',

      createdAt:
        formData.createdAt ||
        Date.now(),
    };

    if (formData.id) {
      setBooks((currentBooks) =>
        currentBooks.map((b) =>
          b.id === formData.id
            ? newBook
            : b
        )
      );
    } else {
      setBooks((currentBooks) => [
        newBook,
        ...currentBooks,
      ]);
    }

    setIsAddModalOpen(false);

    setSelectedBookDetail(null);

    resetForm();
  };

  // ==========================================================
  // RESET
  // ==========================================================

  const resetForm = () => {
    setFormData({
      title: '',
      author: '',
      publishCountry: '',
      coverUrl: '',
      publisher: '',
      publishYear: '',
      pages: undefined,
      genre: '',
      seriesTag: '',
      volume: '',
      isClassic: false,
      format: 'cartaceo',
      isRead: false,
      readMonth: 'Gennaio',
      readYear: currentYearNum,
      rating: 5,
      notes: '',
      isbn: '',
    });

    setIsbnInput('');
  };

  // ==========================================================
  // MODIFICA
  // ==========================================================

  const handleEditBook = (
    book: BookItem
  ) => {
    setFormData(book);

    setIsbnInput(
      book.isbn || ''
    );

    setSelectedBookDetail(null);

    setIsAddModalOpen(true);
  };

  // ==========================================================
  // ELIMINA
  // ==========================================================

  const handleDeleteBook = (
    id: string
  ) => {
    if (
      confirm(
        'Sei sicuro di voler eliminare questo libro?'
      )
    ) {
      setBooks((currentBooks) =>
        currentBooks.filter(
          (b) => b.id !== id
        )
      );

      if (
        selectedBookDetail?.id === id
      ) {
        setSelectedBookDetail(null);
      }
    }
  };

  // ==========================================================
  // DRAG
  // ==========================================================

  const handleDragStart = (
    index: number
  ) => {
    setDraggedIndex(index);
  };

  const handleDragOver = (
    e: React.DragEvent,
    index: number
  ) => {
    e.preventDefault();

    if (
      draggedIndex === null ||
      draggedIndex === index
    ) {
      return;
    }

    const readOnlyBooks =
      books.filter(
        (b) => b.isRead
      );

    const itemToMove =
      readOnlyBooks[draggedIndex];

    if (!itemToMove) return;

    const updated = [...books];

    const sourceGlobalIdx =
      updated.findIndex(
        (b) =>
          b.id === itemToMove.id
      );

    const targetGlobalIdx =
      updated.findIndex(
        (b) =>
          b.id ===
          readOnlyBooks[index].id
      );

    if (
      sourceGlobalIdx === -1 ||
      targetGlobalIdx === -1
    ) {
      return;
    }

    updated.splice(
      sourceGlobalIdx,
      1
    );

    updated.splice(
      targetGlobalIdx,
      0,
      itemToMove
    );

    setDraggedIndex(index);

    setBooks(updated);
  };

  const handleDragEnd = () => {
    setDraggedIndex(null);
  };

  // ==========================================================
  // EXPORT EXCEL
  // ==========================================================

  const exportToExcel = () => {
    const dataToExport =
      books.map((b) => ({
        Titolo: b.title,
        Autore: b.author,
        'Paese di Pubblicazione':
          b.publishCountry || '-',
        Classico:
          b.isClassic ? 'Sì' : 'No',
        Stato:
          b.isRead
            ? 'Letto'
            : 'In Biblioteca',
        Formato:
          b.format === 'cartaceo'
            ? 'Cartaceo'
            : b.format === 'ebook'
            ? 'eBook'
            : 'eBook + Cartaceo',
        Editore:
          b.publisher,
        'Anno Pubblicazione':
          b.publishYear,
        Genere:
          b.genre,
        'Serie / Tag':
          b.seriesTag,
        Volume:
          b.volume || '-',
        Pagine:
          b.pages,
        'Mese e Anno di Lettura':
          b.readMonthYear,
        ISBN:
          b.isbn,
      }));

    const worksheet =
      XLSX.utils.json_to_sheet(
        dataToExport
      );

    const workbook =
      XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(
      workbook,
      worksheet,
      'Biblioteca'
    );

    XLSX.writeFile(
      workbook,
      'La_Mia_Biblioteca.xlsx'
    );
  };

  // ==========================================================
  // EXPORT PDF
  // ==========================================================

  const exportToPDF = () => {
    const doc = new jsPDF();

    doc.text(
      'La Mia Biblioteca - Report',
      14,
      15
    );

    const tableData =
      books.map((b) => [
        b.title,
        b.author,
        b.publishCountry || '-',
        b.isClassic ? 'Sì' : 'No',
        b.isRead
          ? 'Letto'
          : 'In Libreria',
        b.format === 'cartaceo'
          ? 'Cartaceo'
          : b.format === 'ebook'
          ? 'eBook'
          : 'eBook + Cartaceo',
        b.genre || '-',
        b.readMonthYear || '-',
      ]);

    autoTable(doc, {
      head: [
        [
          'Titolo',
          'Autore',
          'Paese',
          'Classico',
          'Stato',
          'Formato',
          'Genere',
          'Data Lettura',
        ],
      ],
      body: tableData,
      startY: 20,
    });

    doc.save(
      'La_Mia_Biblioteca.pdf'
    );
  };

  // ==========================================================
  // BACKUP
  // ==========================================================

  const exportBackup = () => {
    const dataStr =
      'data:text/json;charset=utf-8,' +
      encodeURIComponent(
        JSON.stringify(books)
      );

    const downloadAnchor =
      document.createElement('a');

    downloadAnchor.setAttribute(
      'href',
      dataStr
    );

    downloadAnchor.setAttribute(
      'download',
      'backup_libreria.json'
    );

    document.body.appendChild(
      downloadAnchor
    );

    downloadAnchor.click();

    downloadAnchor.remove();
  };

  // ==========================================================
  // IMPORT BACKUP
  // ==========================================================

  const handleImportBackup = (
    e: React.ChangeEvent<HTMLInputElement>
  ) => {
    const fileReader =
      new FileReader();

    if (
      e.target.files &&
      e.target.files[0]
    ) {
      fileReader.readAsText(
        e.target.files[0],
        'UTF-8'
      );

      fileReader.onload = (
        event
      ) => {
        try {
          const parsed =
            JSON.parse(
              event.target?.result as string
            );

          if (Array.isArray(parsed)) {
            setBooks(parsed);

            alert(
              'Backup ripristinato con successo!'
            );
          }
        } catch {
          alert(
            'File di backup non valido.'
          );
        }
      };
    }
  };

  // ==========================================================
  // CANCELLA TUTTO
  // ==========================================================

  const handleClearAll = () => {
    if (
      confirm(
        'ATTENZIONE: Verranno cancellati TUTTI i libri salvati. Procedere?'
      )
    ) {
      setBooks([]);

      localStorage.removeItem(
        'ios_library_books_v5'
      );
    }
  };

  // ==========================================================
  // FILTRI LIBRI LETTI
  // ==========================================================

  const readBooksFiltered =
    books
      .filter((b) => b.isRead)

      .filter((b) =>
        filterGenre === 'all'
          ? true
          : b.genre === filterGenre
      )

      .filter((b) => {
        if (
          filterFormat === 'all'
        ) {
          return true;
        }

        if (
          filterFormat ===
          'cartaceo'
        ) {
          return (
            b.format ===
            'cartaceo'
          );
        }

        if (
          filterFormat ===
          'ebook'
        ) {
          return (
            b.format ===
              'ebook' ||
            b.format ===
              'ebook_and_paper'
          );
        }

        return true;
      })

      .filter((b) => {
        if (
          filterYear === 'all'
        ) {
          return true;
        }

        return (
          b.readYear?.toString() ===
            filterYear ||
          b.readMonthYear?.includes(
            filterYear
          )
        );
      })

      .filter((b) =>
        searchQuery === ''
          ? true
          : b.title
              .toLowerCase()
              .includes(
                searchQuery.toLowerCase()
              ) ||
            b.author
              .toLowerCase()
              .includes(
                searchQuery.toLowerCase()
              )
      );

  // ==========================================================
  // AUTORI
  // ==========================================================

  const authorsMap =
    books.reduce(
      (acc, book) => {
        const authorName =
          book.author.trim() ||
          'Autore Sconosciuto';

        if (!acc[authorName]) {
          acc[authorName] = [];
        }

        acc[authorName].push(book);

        return acc;
      },
      {} as Record<
        string,
        BookItem[]
      >
    );

  const sortedAuthors =
    Object.keys(
      authorsMap
    ).sort((a, b) =>
      a.localeCompare(b)
    );

  // ==========================================================
  // GENERI
  // ==========================================================

  const genresMap =
    books.reduce(
      (acc, book) => {
        const genreName =
          book.genre?.trim() ||
          'Generico / Altro';

        if (!acc[genreName]) {
          acc[genreName] = [];
        }

        acc[genreName].push(book);

        return acc;
      },
      {} as Record<
        string,
        BookItem[]
      >
    );

  const sortedGenres =
    Object.keys(
      genresMap
    ).sort((a, b) =>
      a.localeCompare(b)
    );

  const availableGenres =
    Array.from(
      new Set(
        books
          .map((b) => b.genre)
          .filter(Boolean)
      )
    );

  const availableYears =
    Array.from(
      new Set(
        books
          .map((b) => {
            if (b.readYear) {
              return b.readYear.toString();
            }

            const match =
              b.readMonthYear?.match(
                /\d{4}/
              );

            return match
              ? match[0]
              : null;
          })
          .filter(Boolean)
      )
    );

  // ==========================================================
  // LABEL FORMATO
  // ==========================================================

  const formatLabel = (
    format: string
  ) => {
    if (
      format === 'cartaceo'
    ) {
      return 'Cartaceo';
    }

    if (
      format === 'ebook'
    ) {
      return 'eBook';
    }

    if (
      format ===
      'ebook_and_paper'
    ) {
      return 'eBook + Cartaceo';
    }

    return format;
  };

  // ==========================================================
  // GRUPPO AUTORI
  // ==========================================================

  const renderAuthorGroup = (
    bookList: BookItem[]
  ) => {
    const map =
      bookList.reduce(
        (acc, book) => {
          const author =
            book.author.trim() ||
            'Autore Sconosciuto';

          if (!acc[author]) {
            acc[author] = [];
          }

          acc[author].push(book);

          return acc;
        },
        {} as Record<
          string,
          BookItem[]
        >
      );

    return Object.keys(map)
      .sort()
      .map((author) => (
        <div
          key={author}
          className="space-y-2 pt-2"
        >
          <h3 className="font-serif font-bold text-sm text-amber-950 border-b border-amber-900/10 pb-1">
            {author}
          </h3>

          <div className="grid grid-cols-2 gap-3">
            {map[author]
              .sort((a, b) =>
                (a.volume || '')
                  .localeCompare(
                    b.volume || ''
                  )
              )
              .map((book) => (
                <div
                  key={book.id}
                  onClick={() =>
                    setSelectedBookDetail(
                      book
                    )
                  }
                  className="bg-[#FFFDF9] p-3 rounded-2xl shadow-sm border border-amber-900/10 flex flex-col cursor-pointer"
                >
                  <div className="w-full h-36 bg-amber-100/40 rounded-xl overflow-hidden relative mb-2 border border-amber-900/10">
                    {book.coverUrl ? (
                      <img
                        src={
                          book.coverUrl
                        }
                        alt={
                          book.title
                        }
                        className="w-full h-full object-cover"
                      />
                    ) : (
                      <div className="w-full h-full flex items-center justify-center text-amber-800/30">
                        <Book className="w-6 h-6" />
                      </div>
                    )}

                    {book.isRead && (
                      <div className="absolute top-2 right-2 bg-emerald-700 text-amber-50 p-1 rounded-full shadow">
                        <CheckCircle2 className="w-3 h-3" />
                      </div>
                    )}

                    {book.volume && (
                      <div className="absolute top-2 left-2 bg-amber-800 text-amber-50 text-[9px] font-bold px-2 py-0.5 rounded-full">
                        Vol.{' '}
                        {book.volume}
                      </div>
                    )}
                  </div>

                  <h4 className="font-serif font-bold text-xs text-amber-950 line-clamp-2">
                    {book.title}
                  </h4>

                  <p className="text-[10px] text-amber-800/60 mt-0.5">
                    {book.publishYear
                      ? `Anno: ${book.publishYear}`
                      : ''}
                  </p>
                </div>
              ))}
          </div>
        </div>
      ));
  };

  // ==========================================================
  // RENDER
  // ==========================================================

  return (
    <div className="min-h-screen bg-[#FBF9F5] text-amber-950 font-sans pb-24 select-none">
      {/* ====================================================
          HEADER
      ==================================================== */}

      <header className="sticky top-0 z-20 bg-[#FBF9F5]/90 backdrop-blur-md border-b border-amber-900/10 px-5 py-3.5 flex justify-between items-center">
        <div>
          <span className="text-[11px] font-bold text-amber-800/70 uppercase tracking-widest block">
            {activeTab === 'home' &&
              'La Mia Collezione'}

            {activeTab === 'read' &&
              'Cronologia Letture'}

            {activeTab ===
              'authors' &&
              'Catalogo Autori'}

            {activeTab ===
              'settings' &&
              'Gestione Dati'}
          </span>

          <h1 className="text-2xl font-serif font-extrabold tracking-tight text-amber-950">
            {activeTab === 'home' &&
              'Home'}

            {activeTab === 'read' &&
              'Libri Letti'}

            {activeTab ===
              'authors' &&
              (selectedAuthor ||
                'Autori')}

            {activeTab ===
              'settings' &&
              'Impostazioni'}
          </h1>
        </div>

        {activeTab === 'read' && (
          <button
            onClick={() => {
              resetForm();
              setIsAddModalOpen(
                true
              );
            }}
            className="w-10 h-10 bg-amber-800 text-amber-50 rounded-full flex items-center justify-center shadow-md active:scale-95 transition-transform"
          >
            <Plus className="w-5 h-5 stroke-[2.5]" />
          </button>
        )}
      </header>

      {/* ====================================================
          HOME
      ==================================================== */}

      {activeTab === 'home' && (
        <div className="p-4 space-y-5 max-w-lg mx-auto">
          {homeSubView !==
          'none' ? (
            <div className="space-y-4">
              <button
                onClick={() => {
                  setHomeSubView(
                    'none'
                  );

                  setSelectedGenreHome(
                    null
                  );
                }}
                className="text-xs font-bold text-amber-800 flex items-center gap-1"
              >
                ← Torna alla Home
              </button>

              {homeSubView ===
                'classics' && (
                <div className="space-y-3">
                  <div className="flex items-center gap-2">
                    <Sparkles className="w-5 h-5 text-amber-600" />

                    <h2 className="text-lg font-serif font-bold text-amber-950">
                      I Miei Classici
                    </h2>
                  </div>

                  {renderAuthorGroup(
                    books.filter(
                      (b) =>
                        b.isClassic
                    )
                  )}
                </div>
              )}

              {homeSubView ===
                'genres' && (
                <div className="space-y-4">
                  {!selectedGenreHome ? (
                    <div className="space-y-2">
                      <h2 className="text-lg font-serif font-bold text-amber-950 mb-2">
                        Generi nella
                        Biblioteca
                      </h2>

                      {sortedGenres.map(
                        (
                          genre
                        ) => (
                          <div
                            key={
                              genre
                            }
                            onClick={() =>
                              setSelectedGenreHome(
                                genre
                              )
                            }
                            className="bg-[#FFFDF9] p-4 rounded-2xl shadow-sm border border-amber-900/10 flex justify-between items-center cursor-pointer"
                          >
                            <div>
                              <h3 className="font-serif font-bold text-base text-amber-950">
                                {
                                  genre
                                }
                              </h3>

                              <p className="text-xs text-amber-800/60 font-medium">
                                {
                                  genresMap[
                                    genre
                                  ]
                                    .length
                                }{' '}
                                {genresMap[
                                  genre
                                ]
                                  .length ===
                                1
                                  ? 'libro'
                                  : 'libri'}
                              </p>
                            </div>

                            <ChevronRight className="w-5 h-5 text-amber-800/30" />
                          </div>
                        )
                      )}
                    </div>
                  ) : (
                    <div className="space-y-3">
                      <button
                        onClick={() =>
                          setSelectedGenreHome(
                            null
                          )
                        }
                        className="text-xs font-bold text-amber-800 mb-1 block"
                      >
                        ← Tutti i
                        generi
                      </button>

                      <h2 className="text-lg font-serif font-bold text-amber-950">
                        {
                          selectedGenreHome
                        }
                      </h2>

                      {renderAuthorGroup(
                        genresMap[
                          selectedGenreHome
                        ]
                      )}
                    </div>
                  )}
                </div>
              )}
            </div>
          ) : (
            <>
              <div className="bg-[#FFFDF9] rounded-3xl p-5 shadow-sm border border-amber-900/10 space-y-3">
                <div className="flex items-center gap-4">
                  <div className="w-12 h-12 rounded-2xl bg-amber-800 text-amber-50 flex items-center justify-center shadow-md flex-shrink-0">
                    <Calendar className="w-6 h-6" />
                  </div>

                  <div className="flex-1">
                    <span className="text-xs font-bold text-amber-800/60 uppercase tracking-wider block">
                      Anno{' '}
                      {currentYearNum}
                    </span>

                    <span className="text-2xl font-serif font-bold text-amber-950">
                      {
                        readThisYearCount
                      }{' '}
                      <span className="text-sm font-sans font-normal text-amber-800/70">
                        libri letti
                      </span>
                    </span>
                  </div>
                </div>

                <div className="pt-2.5 border-t border-amber-900/10 flex items-center justify-center text-xs font-semibold text-amber-900/80">
                  <span>
                    {
                      readThisYearCartacei
                    }{' '}
                    Cartacei e{' '}
                    {
                      readThisYearEbook
                    }{' '}
                    eBook
                  </span>
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3.5">
                <div className="bg-[#FFFDF9] p-4 rounded-3xl shadow-sm border border-amber-900/10 flex flex-col justify-between min-h-[9rem]">
                  <div className="w-10 h-10 rounded-2xl bg-amber-100/70 text-amber-900 flex items-center justify-center">
                    <BookOpen className="w-5 h-5" />
                  </div>

                  <div className="mt-2">
                    <span className="text-3xl font-serif font-black text-amber-950 leading-none">
                      {totalBooks}
                    </span>

                    <span className="text-xs font-bold text-amber-800/60 uppercase tracking-wider block mt-1">
                      In Biblioteca
                    </span>

                    <div className="mt-2 pt-2 border-t border-amber-900/10 text-xs font-semibold text-amber-900/80">
                      {
                        totalCartacei
                      }{' '}
                      Cartacei e{' '}
                      {totalEbook}{' '}
                      eBook
                    </div>
                  </div>
                </div>

                <div className="bg-[#FFFDF9] p-4 rounded-3xl shadow-sm border border-amber-900/10 flex flex-col justify-between min-h-[9rem]">
                  <div className="w-10 h-10 rounded-2xl bg-emerald-100/70 text-emerald-800 flex items-center justify-center">
                    <CheckCircle2 className="w-5 h-5" />
                  </div>

                  <div className="mt-2">
                    <span className="text-3xl font-serif font-black text-amber-950 leading-none">
                      {
                        readBooksCount
                      }
                    </span>

                    <span className="text-xs font-bold text-amber-800/60 uppercase tracking-wider block mt-1">
                      Libri Letti
                    </span>

                    <div className="mt-2 pt-2 border-t border-amber-900/10 text-xs font-semibold text-amber-900/80">
                      {
                        readCartacei
                      }{' '}
                      Cartacei e{' '}
                      {readEbook}{' '}
                      eBook
                    </div>
                  </div>
                </div>
              </div>

              <div className="space-y-3 pt-2">
                <h2 className="text-xs font-bold text-amber-800/60 uppercase tracking-wider px-1">
                  Esplora Categorie
                </h2>

                <div className="grid grid-cols-2 gap-3.5">
                  <div
                    onClick={() =>
                      setHomeSubView(
                        'classics'
                      )
                    }
                    className="bg-gradient-to-br from-amber-700 to-amber-900 text-amber-50 p-4 rounded-3xl shadow-md cursor-pointer active:scale-95 transition-transform flex flex-col justify-between h-32 relative overflow-hidden"
                  >
                    <Sparkles className="w-6 h-6 text-amber-200" />

                    <div>
                      <span className="text-lg font-serif font-bold block">
                        Classici
                      </span>

                      <span className="text-xs text-amber-200/80 font-medium">
                        {
                          books.filter(
                            (b) =>
                              b.isClassic
                          ).length
                        }{' '}
                        libri
                        conservati
                      </span>
                    </div>
                  </div>

                  <div
                    onClick={() =>
                      setHomeSubView(
                        'genres'
                      )
                    }
                    className="bg-gradient-to-br from-stone-800 to-amber-950 text-amber-50 p-4 rounded-3xl shadow-md cursor-pointer active:scale-95 transition-transform flex flex-col justify-between h-32 relative overflow-hidden"
                  >
                    <Layers className="w-6 h-6 text-amber-300" />

                    <div>
                      <span className="text-lg font-serif font-bold block">
                        Generi
                      </span>

                      <span className="text-xs text-amber-200/80 font-medium">
                        {
                          sortedGenres.length
                        }{' '}
                        categorie
                      </span>
                    </div>
                  </div>
                </div>
              </div>
            </>
          )}
        </div>
      )}

      {/* ====================================================
          LIBRI LETTI
      ==================================================== */}

      {activeTab === 'read' && (
        <div className="p-4 space-y-4 max-w-lg mx-auto">
          <div className="space-y-2.5">
            <div className="relative">
              <Search className="w-4 h-4 absolute left-3.5 top-3 text-amber-800/40" />

              <input
                type="text"
                placeholder="Cerca nei libri letti..."
                value={searchQuery}
                onChange={(e) =>
                  setSearchQuery(
                    e.target.value
                  )
                }
                className="w-full pl-9 pr-4 py-2.5 bg-[#FFFDF9] border border-amber-900/10 rounded-2xl text-xs font-medium focus:outline-none focus:ring-2 focus:ring-amber-800/20"
              />
            </div>

            <div className="flex gap-2 overflow-x-auto pb-1 text-xs">
              <select
                value={
                  filterGenre
                }
                onChange={(e) =>
                  setFilterGenre(
                    e.target.value
                  )
                }
                className="bg-[#FFFDF9] border border-amber-900/10 rounded-xl px-3 py-2 font-medium text-amber-950"
              >
                <option value="all">
                  Tutti i generi
                </option>

                {availableGenres.map(
                  (g) => (
                    <option
                      key={g}
                      value={g}
                    >
                      {g}
                    </option>
                  )
                )}
              </select>

              <select
                value={
                  filterFormat
                }
                onChange={(e) =>
                  setFilterFormat(
                    e.target.value
                  )
                }
                className="bg-[#FFFDF9] border border-amber-900/10 rounded-xl px-3 py-2 font-medium text-amber-950"
              >
                <option value="all">
                  Tutti i formati
                </option>

                <option value="cartaceo">
                  Cartaceo
                </option>

                <option value="ebook">
                  eBook
                </option>
              </select>

              <select
                value={filterYear}
                onChange={(e) =>
                  setFilterYear(
                    e.target.value
                  )
                }
                className="bg-[#FFFDF9] border border-amber-900/10 rounded-xl px-3 py-2 font-medium text-amber-950"
              >
                <option value="all">
                  Tutti gli anni
                </option>

                {availableYears.map(
                  (y) => (
                    <option
                      key={y}
                      value={y!}
                    >
                      {y}
                    </option>
                  )
                )}
              </select>
            </div>
          </div>

          <p className="text-[11px] text-amber-800/50 font-medium px-1">
            Trascina tramite
            l'icona a sinistra
            per riordinare
            l'elenco dei libri
            letti.
          </p>

          <div className="space-y-3">
            {readBooksFiltered.length ===
            0 ? (
              <div className="text-center py-12 bg-[#FFFDF9] rounded-3xl border border-dashed border-amber-900/20">
                <BookOpen className="w-8 h-8 text-amber-800/30 mx-auto mb-2" />

                <p className="text-xs text-amber-800/60 font-medium">
                  Nessun libro
                  letto
                  trovato.
                </p>
              </div>
            ) : (
              readBooksFiltered.map(
                (
                  book,
                  index
                ) => (
                  <div
                    key={
                      book.id
                    }
                    draggable
                    onDragStart={() =>
                      handleDragStart(
                        index
                      )
                    }
                    onDragOver={(
                      e
                    ) =>
                      handleDragOver(
                        e,
                        index
                      )
                    }
                    onDragEnd={
                      handleDragEnd
                    }
                    className={`bg-[#FFFDF9] p-3 rounded-2xl shadow-sm border border-amber-900/10 flex items-center gap-3 transition-colors ${
                      draggedIndex ===
                      index
                        ? 'bg-amber-100/50'
                        : 'hover:bg-amber-50/50'
                    }`}
                  >
                    <div className="cursor-grab active:cursor-grabbing text-amber-800/30 hover:text-amber-800">
                      <GripVertical className="w-5 h-5" />
                    </div>

                    <div
                      onClick={() =>
                        setSelectedBookDetail(
                          book
                        )
                      }
                      className="w-13 h-19 bg-amber-100/40 rounded-lg overflow-hidden flex-shrink-0 relative border border-amber-900/10 cursor-pointer"
                    >
                      {book.coverUrl ? (
                        <img
                          src={
                            book.coverUrl
                          }
                          alt={
                            book.title
                          }
                          className="w-full h-full object-cover"
                        />
                      ) : (
                        <div className="w-full h-full flex items-center justify-center text-amber-800/30">
                          <Book className="w-5 h-5" />
                        </div>
                      )}
                    </div>

                    <div
                      onClick={() =>
                        setSelectedBookDetail(
                          book
                        )
                      }
                      className="flex-1 min-w-0 cursor-pointer"
                    >
                      <h3 className="font-serif font-bold text-sm text-amber-950 truncate">
                        {
                          book.title
                        }
                      </h3>

                      <p className="text-xs text-amber-800/70 truncate">
                        {
                          book.author
                        }
                      </p>

                      {book.rating && (
                        <div className="flex text-amber-500 mt-1">
                          {[
                            ...Array(
                              5
                            ),
                          ].map(
                            (
                              _,
                              i
                            ) => (
                              <Star
                                key={
                                  i
                                }
                                className={`w-3 h-3 ${
                                  i <
                                  book.rating!
                                    ? 'fill-amber-500'
                                    : 'text-amber-200'
                                }`}
                              />
                            )
                          )}
                        </div>
                      )}

                      <p className="text-[10px] font-semibold text-amber-800/50 mt-1">
                        Letto:{' '}
                        <span className="text-amber-950">
                          {book.readMonthYear ||
                            'Data non specificata'}
                        </span>
                      </p>
                    </div>

                    <ChevronRight className="w-4 h-4 text-amber-800/30" />
                  </div>
                )
              )
            )}
          </div>
        </div>
      )}

      {/* ====================================================
          AUTORI
      ==================================================== */}

      {activeTab === 'authors' && (
        <div className="p-4 space-y-4 max-w-lg mx-auto">
          {!selectedAuthor ? (
            <div className="space-y-2">
              {sortedAuthors.length ===
              0 ? (
                <div className="text-center py-12 bg-[#FFFDF9] rounded-3xl border border-dashed border-amber-900/20">
                  <Users className="w-8 h-8 text-amber-800/30 mx-auto mb-2" />

                  <p className="text-xs text-amber-800/60 font-medium">
                    Nessun autore
                    presente.
                  </p>
                </div>
              ) : (
                sortedAuthors.map(
                  (author) => {
                    const authorBooks =
                      authorsMap[
                        author
                      ];

                    return (
                      <div
                        key={
                          author
                        }
                        onClick={() =>
                          setSelectedAuthor(
                            author
                          )
                        }
                        className="bg-[#FFFDF9] p-4 rounded-2xl shadow-sm border border-amber-900/10 flex justify-between items-center cursor-pointer"
                      >
                        <div>
                          <h3 className="font-serif font-bold text-base text-amber-950">
                            {
                              author
                            }
                          </h3>

                          <p className="text-xs text-amber-800/60 font-medium">
                            {
                              authorBooks.length
                            }{' '}
                            {authorBooks.length ===
                            1
                              ? 'libro'
                              : 'libri'}
                          </p>
                        </div>

                        <ChevronRight className="w-5 h-5 text-amber-800/30" />
                      </div>
                    );
                  }
                )
              )}
            </div>
          ) : (
            <div className="space-y-4">
              <button
                onClick={() =>
                  setSelectedAuthor(
                    null
                  )
                }
                className="text-xs font-bold text-amber-800 flex items-center gap-1 mb-2"
              >
                ← Torna alla lista
                autori
              </button>

              <div className="grid grid-cols-2 gap-3.5">
                {authorsMap[
                  selectedAuthor
                ]?.map((book) => (
                  <div
                    key={
                      book.id
                    }
                    onClick={() =>
                      setSelectedBookDetail(
                        book
                      )
                    }
                    className="bg-[#FFFDF9] p-3 rounded-2xl shadow-sm border border-amber-900/10 flex flex-col cursor-pointer relative"
                  >
                    <div className="w-full h-44 bg-amber-100/40 rounded-xl overflow-hidden relative border border-amber-900/10 mb-2">
                      {book.coverUrl ? (
                        <img
                          src={
                            book.coverUrl
                          }
                          alt={
                            book.title
                          }
                          className="w-full h-full object-cover"
                        />
                      ) : (
                        <div className="w-full h-full flex items-center justify-center text-amber-800/30">
                          <Book className="w-8 h-8" />
                        </div>
                      )}

                      {book.isRead && (
                        <div className="absolute top-2 right-2 bg-emerald-700 text-amber-50 p-1 rounded-full shadow-md">
                          <CheckCircle2 className="w-3.5 h-3.5" />
                        </div>
                      )}

                      <div className="absolute bottom-2 left-2 bg-amber-950/70 backdrop-blur-md text-amber-50 text-[9px] font-medium px-2 py-0.5 rounded-full">
                        {formatLabel(
                          book.format
                        )}
                      </div>

                      {book.volume && (
                        <div className="absolute top-2 left-2 bg-amber-800 text-amber-50 text-[9px] font-bold px-2 py-0.5 rounded-full shadow">
                          Vol.{' '}
                          {
                            book.volume
                          }
                        </div>
                      )}
                    </div>

                    <h4 className="font-serif font-bold text-xs text-amber-950 line-clamp-2">
                      {
                        book.title
                      }
                    </h4>

                    <p className="text-[10px] font-medium text-amber-800/60 mt-1">
                      {book.publishYear
                        ? `Anno: ${book.publishYear}`
                        : ''}
                    </p>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      )}

      {/* ====================================================
          SETTINGS
      ==================================================== */}

      {activeTab ===
        'settings' && (
        <div className="p-4 space-y-5 max-w-lg mx-auto">
          <div className="bg-[#FFFDF9] rounded-3xl p-4 shadow-sm border border-amber-900/10 space-y-3">
            <h2 className="text-xs font-bold text-amber-800/60 uppercase tracking-wider">
              Esporta la tua
              Biblioteca
            </h2>

            <button
              onClick={
                exportToExcel
              }
              className="w-full p-3 bg-emerald-50 text-emerald-900 rounded-2xl font-bold text-xs flex items-center gap-3 border border-emerald-200"
            >
              <FileSpreadsheet className="w-5 h-5 text-emerald-700" />

              Esporta in Foglio
              Excel (.xlsx)
            </button>

            <button
              onClick={
                exportToPDF
              }
              className="w-full p-3 bg-rose-50 text-rose-900 rounded-2xl font-bold text-xs flex items-center gap-3 border border-rose-200"
            >
              <FileText className="w-5 h-5 text-rose-700" />

              Esporta Report PDF
            </button>
          </div>

          <div className="bg-[#FFFDF9] rounded-3xl p-4 shadow-sm border border-amber-900/10 space-y-3">
            <h2 className="text-xs font-bold text-amber-800/60 uppercase tracking-wider">
              Backup &
              Ripristino Dati
            </h2>

            <button
              onClick={
                exportBackup
              }
              className="w-full p-3 bg-amber-100/60 text-amber-950 rounded-2xl font-bold text-xs flex items-center gap-3 border border-amber-200"
            >
              <Download className="w-5 h-5 text-amber-800" />

              Salva Backup Dati
              (JSON)
            </button>

            <label className="w-full p-3 bg-stone-100/80 text-stone-900 rounded-2xl font-bold text-xs flex items-center gap-3 border border-stone-200 cursor-pointer">
              <Upload className="w-5 h-5 text-stone-700" />

              Ripristina Backup
              da File

              <input
                type="file"
                accept=".json"
                onChange={
                  handleImportBackup
                }
                className="hidden"
              />
            </label>
          </div>

          <div className="bg-[#FFFDF9] rounded-3xl p-4 shadow-sm border border-amber-900/10 space-y-3">
            <h2 className="text-xs font-bold text-rose-700 uppercase tracking-wider">
              Zona Pericolo
            </h2>

            <button
              onClick={
                handleClearAll
              }
              className="w-full p-3 bg-rose-700 text-amber-50 rounded-2xl font-bold text-xs flex items-center justify-center gap-2"
            >
              <Trash2 className="w-4 h-4" />

              Cancella Intera
              Biblioteca
            </button>
          </div>
        </div>
      )}

      {/* ====================================================
          DETTAGLIO LIBRO
      ==================================================== */}

      {selectedBookDetail && (
        <div className="fixed inset-0 z-50 bg-amber-950/40 backdrop-blur-sm flex items-end sm:items-center justify-center p-0 sm:p-4">
          <div className="bg-[#FFFDF9] w-full max-w-lg rounded-t-3xl sm:rounded-3xl max-h-[90vh] overflow-y-auto p-6 space-y-4 shadow-2xl relative border border-amber-900/10">
            <button
              onClick={() =>
                setSelectedBookDetail(
                  null
                )
              }
              className="absolute top-4 right-4 p-2 bg-amber-100/50 rounded-full text-amber-900"
            >
              <X className="w-5 h-5" />
            </button>

            <div className="flex gap-4 items-start pt-2">
              <div className="w-24 h-36 bg-amber-100/40 rounded-xl overflow-hidden border border-amber-900/10 flex-shrink-0">
                {selectedBookDetail.coverUrl ? (
                  <img
                    src={
                      selectedBookDetail.coverUrl
                    }
                    alt={
                      selectedBookDetail.title
                    }
                    className="w-full h-full object-cover"
                  />
                ) : (
                  <div className="w-full h-full flex items-center justify-center text-amber-800/30">
                    <Book className="w-8 h-8" />
                  </div>
                )}
              </div>

              <div className="flex-1">
                <h2 className="text-lg font-serif font-bold text-amber-950 leading-tight">
                  {
                    selectedBookDetail.title
                  }
                </h2>

                <p className="text-xs font-medium text-amber-800/70 mt-1">
                  {
                    selectedBookDetail.author
                  }
                </p>

                <div className="flex flex-wrap gap-2 mt-3">
                  <span className="text-[10px] px-2.5 py-1 rounded-full bg-amber-100/70 text-amber-900 font-bold">
                    {formatLabel(
                      selectedBookDetail.format
                    )}
                  </span>

                  {selectedBookDetail.isClassic && (
                    <span className="text-[10px] px-2.5 py-1 rounded-full bg-amber-200/60 text-amber-950 font-bold">
                      Classico
                    </span>
                  )}

                  {selectedBookDetail.isRead ? (
                    <span className="text-[10px] px-2.5 py-1 rounded-full bg-emerald-100/70 text-emerald-900 font-bold">
                      Letto
                    </span>
                  ) : (
                    <span className="text-[10px] px-2.5 py-1 rounded-full bg-stone-200/70 text-stone-900 font-bold">
                      In Biblioteca
                    </span>
                  )}
                </div>
              </div>
            </div>

            <div className="grid grid-cols-2 gap-3 bg-amber-50/60 p-4 rounded-2xl text-xs text-amber-900">
              <div>
                <span className="block text-amber-800/50 font-semibold text-[10px]">
                  Paese di
                  Pubblicazione:
                </span>

                <span className="font-bold">
                  {selectedBookDetail.publishCountry ||
                    '-'}
                </span>
              </div>

              <div>
                <span className="block text-amber-800/50 font-semibold text-[10px]">
                  Editore:
                </span>

                <span className="font-bold">
                  {selectedBookDetail.publisher ||
                    '-'}
                </span>
              </div>

              <div>
                <span className="block text-amber-800/50 font-semibold text-[10px]">
                  Anno
                  Pubblicazione:
                </span>

                <span className="font-bold">
                  {selectedBookDetail.publishYear ||
                    '-'}
                </span>
              </div>

              <div>
                <span className="block text-amber-800/50 font-semibold text-[10px]">
                  Pagine:
                </span>

                <span className="font-bold">
                  {selectedBookDetail.pages ||
                    '-'}
                </span>
              </div>

              <div>
                <span className="block text-amber-800/50 font-semibold text-[10px]">
                  Genere:
                </span>

                <span className="font-bold">
                  {selectedBookDetail.genre ||
                    '-'}
                </span>
              </div>

              <div>
                <span className="block text-amber-800/50 font-semibold text-[10px]">
                  Serie / Tag:
                </span>

                <span className="font-bold">
                  {selectedBookDetail.seriesTag ||
                    '-'}
                </span>
              </div>

              <div>
                <span className="block text-amber-800/50 font-semibold text-[10px]">
                  Volume:
                </span>

                <span className="font-bold">
                  {selectedBookDetail.volume ||
                    '-'}
                </span>
              </div>

              <div>
                <span className="block text-amber-800/50 font-semibold text-[10px]">
                  ISBN:
                </span>

                <span className="font-bold break-all">
                  {selectedBookDetail.isbn ||
                    '-'}
                </span>
              </div>

              {selectedBookDetail.isRead && (
                <div className="col-span-2 border-t border-amber-900/10 pt-2">
                  <span className="block text-amber-800/50 font-semibold text-[10px]">
                    Mese & Anno di
                    Lettura:
                  </span>

                  <span className="font-bold">
                    {selectedBookDetail.readMonthYear ||
                      '-'}
                  </span>
                </div>
              )}
            </div>

            <div className="grid grid-cols-2 gap-3 pt-2">
              <button
                onClick={() =>
                  handleEditBook(
                    selectedBookDetail
                  )
                }
                className="py-3 bg-amber-800 text-amber-50 rounded-xl font-bold text-xs flex items-center justify-center gap-2"
              >
                <Edit className="w-4 h-4" />

                Modifica
              </button>

              <button
                onClick={() =>
                  handleDeleteBook(
                    selectedBookDetail.id
                  )
                }
                className="py-3 bg-rose-100 text-rose-800 rounded-xl font-bold text-xs flex items-center justify-center gap-2 border border-rose-200"
              >
                <Trash2 className="w-4 h-4" />

                Elimina
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ====================================================
          MODALE AGGIUNTA / MODIFICA
      ==================================================== */}

      {isAddModalOpen && (
        <div className="fixed inset-0 z-50 bg-amber-950/40 backdrop-blur-sm flex items-end sm:items-center justify-center p-0 sm:p-4">
          <div className="bg-[#FFFDF9] w-full max-w-lg rounded-t-3xl sm:rounded-3xl max-h-[92vh] overflow-y-auto p-6 space-y-4 shadow-2xl relative border border-amber-900/10">
            {/* HEADER */}

            <div className="flex justify-between items-center border-b border-amber-900/10 pb-3">
              <h2 className="text-base font-serif font-bold text-amber-950">
                {formData.id
                  ? 'Modifica Libro'
                  : 'Aggiungi Nuovo Libro'}
              </h2>

              <button
                onClick={() =>
                  setIsAddModalOpen(
                    false
                  )
                }
                className="p-2 text-amber-800/40"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* =================================================
                SCANNER ISBN
            ================================================== */}

            <div className="bg-amber-100/50 p-3.5 rounded-2xl space-y-3 border border-amber-900/10">
              <div className="flex justify-between items-center">
                <div>
                  <span className="text-[10px] font-bold text-amber-900 uppercase tracking-wider block">
                    Inserimento ISBN
                  </span>

                  <span className="text-[10px] text-amber-800/60">
                    Scansiona il
                    codice o
                    inseriscilo
                    manualmente
                  </span>
                </div>

                <Scan className="w-4 h-4 text-amber-800" />
              </div>

              {/* PULSANTE FOTOCAMERA */}

              <button
                type="button"
                onClick={() =>
                  setIsScannerOpen(
                    true
                  )
                }
                className="w-full py-3 bg-amber-800 text-amber-50 rounded-xl text-xs font-bold flex items-center justify-center gap-2 shadow-md active:scale-[0.98] transition-transform"
              >
                <Camera className="w-4 h-4" />

                Scansiona ISBN con
                fotocamera
              </button>

              <div className="flex items-center gap-2">
                <div className="h-px bg-amber-900/10 flex-1" />

                <span className="text-[9px] text-amber-800/40 font-bold uppercase">
                  oppure
                </span>

                <div className="h-px bg-amber-900/10 flex-1" />
              </div>

              {/* INPUT MANUALE */}

              <div className="flex gap-2">
                <input
                  type="text"
                  inputMode="numeric"
                  autoComplete="off"
                  placeholder="ISBN es. 9788806221843"
                  value={
                    isbnInput
                  }
                  onChange={(e) => {
                    const value =
                      e.target.value.replace(
                        /[^0-9Xx-]/g,
                        ''
                      );

                    setIsbnInput(
                      value
                    );

                    setFormData(
                      (
                        prev
                      ) => ({
                        ...prev,
                        isbn: value.replace(
                          /-/g,
                          ''
                        ),
                      })
                    );
                  }}
                  onKeyDown={(
                    e
                  ) => {
                    if (
                      e.key ===
                      'Enter'
                    ) {
                      e.preventDefault();

                      handleSearchBookByISBN();
                    }
                  }}
                  className="flex-1 min-w-0 p-2.5 bg-[#FFFDF9] border border-amber-900/10 rounded-xl text-xs font-medium focus:outline-none focus:ring-2 focus:ring-amber-800/20"
                />

                <button
                  type="button"
                  onClick={() =>
                    handleSearchBookByISBN()
                  }
                  disabled={
                    isSearchingIsbn
                  }
                  className="px-4 bg-stone-800 text-amber-50 rounded-xl text-xs font-bold active:scale-95 transition-transform disabled:opacity-50 flex items-center justify-center min-w-[70px]"
                >
                  {isSearchingIsbn ? (
                    <Loader2 className="w-4 h-4 animate-spin" />
                  ) : (
                    'Cerca'
                  )}
                </button>
              </div>

              {formData.isbn && (
                <div className="flex items-center gap-2 text-[10px] text-emerald-700 font-semibold">
                  <CheckCircle2 className="w-3.5 h-3.5" />

                  ISBN:
                  {
                    formData.isbn
                  }
                </div>
              )}
            </div>

            {/* =================================================
                FORM
            ================================================== */}

            <form
              onSubmit={
                handleSaveBook
              }
              className="space-y-3 text-xs"
            >
              {/* TITOLO */}

              <input
                type="text"
                placeholder="Titolo *"
                value={
                  formData.title ||
                  ''
                }
                onChange={(e) =>
                  setFormData({
                    ...formData,
                    title: e.target
                      .value,
                  })
                }
                className="w-full p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl font-medium focus:outline-none focus:ring-2 focus:ring-amber-800/20"
                required
              />

              {/* AUTORE / PAESE */}

              <div className="grid grid-cols-2 gap-2">
                <input
                  type="text"
                  placeholder="Autore *"
                  value={
                    formData.author ||
                    ''
                  }
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      author:
                        e.target
                          .value,
                    })
                  }
                  className="p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl font-medium focus:outline-none focus:ring-2 focus:ring-amber-800/20"
                  required
                />

                <input
                  type="text"
                  placeholder="Paese di pubblicazione"
                  value={
                    formData.publishCountry ||
                    ''
                  }
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      publishCountry:
                        e.target
                          .value,
                    })
                  }
                  className="p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl font-medium focus:outline-none focus:ring-2 focus:ring-amber-800/20"
                />
              </div>

              {/* COPERTINA */}

              <input
                type="text"
                placeholder="URL Copertina Immagine"
                value={
                  formData.coverUrl ||
                  ''
                }
                onChange={(e) =>
                  setFormData({
                    ...formData,
                    coverUrl:
                      e.target
                        .value,
                  })
                }
                className="w-full p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl font-medium focus:outline-none focus:ring-2 focus:ring-amber-800/20"
              />

              {/* EDITORE / ANNO */}

              <div className="grid grid-cols-2 gap-2">
                <input
                  type="text"
                  placeholder="Editore"
                  value={
                    formData.publisher ||
                    ''
                  }
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      publisher:
                        e.target
                          .value,
                    })
                  }
                  className="p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl font-medium focus:outline-none focus:ring-2 focus:ring-amber-800/20"
                />

                <input
                  type="text"
                  inputMode="numeric"
                  placeholder="Anno Pubblicazione"
                  value={
                    formData.publishYear ||
                    ''
                  }
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      publishYear:
                        e.target
                          .value.replace(
                            /\D/g,
                            ''
                          ),
                    })
                  }
                  className="p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl font-medium focus:outline-none focus:ring-2 focus:ring-amber-800/20"
                />
              </div>

              {/* GENERE / PAGINE */}

              <div className="grid grid-cols-2 gap-2">
                <input
                  type="text"
                  placeholder="Genere"
                  value={
                    formData.genre ||
                    ''
                  }
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      genre:
                        e.target
                          .value,
                    })
                  }
                  className="p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl font-medium focus:outline-none focus:ring-2 focus:ring-amber-800/20"
                />

                <input
                  type="text"
                  inputMode="numeric"
                  pattern="[0-9]*"
                  placeholder="Numero Pagine"
                  value={
                    formData.pages ||
                    ''
                  }
                  onChange={(e) => {
                    const val =
                      e.target.value.replace(
                        /\D/g,
                        ''
                      );

                    setFormData({
                      ...formData,
                      pages: val
                        ? parseInt(
                            val
                          )
                        : undefined,
                    });
                  }}
                  className="p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl font-medium focus:outline-none focus:ring-2 focus:ring-amber-800/20"
                />
              </div>

              {/* SERIE / VOLUME */}

              <div className="grid grid-cols-2 gap-2">
                <input
                  type="text"
                  placeholder="Serie / Tag"
                  value={
                    formData.seriesTag ||
                    ''
                  }
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      seriesTag:
                        e.target
                          .value,
                    })
                  }
                  className="p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl font-medium focus:outline-none focus:ring-2 focus:ring-amber-800/20"
                />

                <input
                  type="text"
                  placeholder="Volume"
                  value={
                    formData.volume ||
                    ''
                  }
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      volume:
                        e.target
                          .value,
                    })
                  }
                  className="p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl font-medium focus:outline-none focus:ring-2 focus:ring-amber-800/20"
                />
              </div>

              {/* CLASSICO */}

              <div className="flex items-center justify-between p-3 bg-amber-50/60 rounded-xl border border-amber-900/10">
                <span className="font-bold text-amber-950 flex items-center gap-1.5">
                  <Sparkles className="w-4 h-4 text-amber-600" />

                  È un Classico?
                </span>

                <input
                  type="checkbox"
                  checked={
                    !!formData.isClassic
                  }
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      isClassic:
                        e.target
                          .checked,
                    })
                  }
                  className="w-5 h-5 accent-amber-800"
                />
              </div>

              {/* FORMATO */}

              <div className="p-3 bg-amber-50/60 rounded-xl space-y-2 border border-amber-900/10">
                <span className="font-bold text-amber-950 block">
                  Formato:
                </span>

                <div className="flex gap-2">
                  <button
                    type="button"
                    onClick={() =>
                      setFormData({
                        ...formData,
                        format:
                          'cartaceo',
                      })
                    }
                    className={`flex-1 py-2 rounded-lg text-xs font-bold transition-colors ${
                      formData.format ===
                      'cartaceo'
                        ? 'bg-amber-800 text-amber-50'
                        : 'bg-[#FFFDF9] border border-amber-900/10 text-amber-900'
                    }`}
                  >
                    Cartaceo
                  </button>

                  <button
                    type="button"
                    onClick={() =>
                      setFormData({
                        ...formData,
                        format:
                          'ebook',
                      })
                    }
                    className={`flex-1 py-2 rounded-lg text-xs font-bold transition-colors ${
                      formData.format ===
                        'ebook' ||
                      formData.format ===
                        'ebook_and_paper'
                        ? 'bg-amber-800 text-amber-50'
                        : 'bg-[#FFFDF9] border border-amber-900/10 text-amber-900'
                    }`}
                  >
                    eBook
                  </button>
                </div>

                {(formData.format ===
                  'ebook' ||
                  formData.format ===
                    'ebook_and_paper') && (
                  <label className="flex items-center gap-2 pt-1 cursor-pointer">
                    <input
                      type="checkbox"
                      checked={
                        formData.format ===
                        'ebook_and_paper'
                      }
                      onChange={(e) =>
                        setFormData({
                          ...formData,
                          format:
                            e.target
                              .checked
                              ? 'ebook_and_paper'
                              : 'ebook',
                        })
                      }
                      className="w-4 h-4 accent-amber-800"
                    />

                    <span className="text-[11px] font-medium text-amber-900">
                      Acquistato anche
                      in formato
                      Cartaceo
                    </span>
                  </label>
                )}
              </div>

              {/* LETTO */}

              <div className="flex items-center justify-between p-3 bg-amber-50/60 rounded-xl border border-amber-900/10">
                <span className="font-bold text-amber-950">
                  Letto
                </span>

                <input
                  type="checkbox"
                  checked={
                    !!formData.isRead
                  }
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      isRead:
                        e.target
                          .checked,
                    })
                  }
                  className="w-5 h-5 accent-amber-800"
                />
              </div>

              {/* DATA LETTURA */}

              {formData.isRead && (
                <div className="p-3 bg-amber-100/40 border border-amber-900/10 rounded-xl space-y-2">
                  <span className="font-bold text-amber-950 block text-[11px] uppercase tracking-wider">
                    Data di Lettura
                  </span>

                  <div className="grid grid-cols-2 gap-2">
                    <div>
                      <label className="block text-[10px] text-amber-800/70 font-semibold mb-1">
                        Mese:
                      </label>

                      <select
                        value={
                          formData.readMonth ||
                          'Gennaio'
                        }
                        onChange={(
                          e
                        ) =>
                          setFormData({
                            ...formData,
                            readMonth:
                              e
                                .target
                                .value,
                          })
                        }
                        className="w-full p-2.5 bg-[#FFFDF9] border border-amber-900/10 rounded-xl font-semibold text-amber-950 focus:outline-none"
                      >
                        {MONTHS.map(
                          (
                            m
                          ) => (
                            <option
                              key={
                                m
                              }
                              value={
                                m
                              }
                            >
                              {
                                m
                              }
                            </option>
                          )
                        )}
                      </select>
                    </div>

                    <div>
                      <label className="block text-[10px] text-amber-800/70 font-semibold mb-1">
                        Anno:
                      </label>

                      <select
                        value={
                          formData.readYear ||
                          currentYearNum
                        }
                        onChange={(
                          e
                        ) =>
                          setFormData({
                            ...formData,
                            readYear:
                              parseInt(
                                e
                                  .target
                                  .value
                              ),
                          })
                        }
                        className="w-full p-2.5 bg-[#FFFDF9] border border-amber-900/10 rounded-xl font-semibold text-amber-950 focus:outline-none"
                      >
                        {yearsList.map(
                          (
                            y
                          ) => (
                            <option
                              key={
                                y
                              }
                              value={
                                y
                              }
                            >
                              {
                                y
                              }
                            </option>
                          )
                        )}
                      </select>
                    </div>
                  </div>
                </div>
              )}

              {/* SALVA */}

              <button
                type="submit"
                className="w-full py-3 bg-amber-800 text-amber-50 rounded-xl font-bold text-xs shadow-md active:scale-95 transition-transform"
              >
                {formData.id
                  ? 'Aggiorna Libro'
                  : 'Salva Libro'}
              </button>
            </form>
          </div>
        </div>
      )}

      {/* ====================================================
          SCANNER CAMERA
      ==================================================== */}

      {isScannerOpen && (
        <ISBNScanner
          onDetected={
            handleISBNDetected
          }
          onClose={() =>
            setIsScannerOpen(
              false
            )
          }
        />
      )}

      {/* ====================================================
          TAB BAR
      ==================================================== */}

      <nav className="fixed bottom-0 left-0 right-0 z-30 bg-[#FBF9F5]/90 backdrop-blur-md border-t border-amber-900/10 flex justify-around py-2.5 max-w-lg mx-auto">
        <button
          onClick={() => {
            setActiveTab(
              'home'
            );

            setSelectedAuthor(
              null
            );

            setHomeSubView(
              'none'
            );
          }}
          className={`flex flex-col items-center gap-1 ${
            activeTab === 'home'
              ? 'text-amber-800'
              : 'text-amber-900/40'
          }`}
        >
          <Home className="w-5 h-5" />

          <span className="text-[10px] font-bold">
            Home
          </span>
        </button>

        <button
          onClick={() => {
            setActiveTab(
              'read'
            );

            setSelectedAuthor(
              null
            );
          }}
          className={`flex flex-col items-center gap-1 ${
            activeTab === 'read'
              ? 'text-amber-800'
              : 'text-amber-900/40'
          }`}
        >
          <BookOpen className="w-5 h-5" />

          <span className="text-[10px] font-bold">
            Letti
          </span>
        </button>

        <button
          onClick={() => {
            setActiveTab(
              'authors'
            );

            setSelectedAuthor(
              null
            );
          }}
          className={`flex flex-col items-center gap-1 ${
            activeTab ===
            'authors'
              ? 'text-amber-800'
              : 'text-amber-900/40'
          }`}
        >
          <Users className="w-5 h-5" />

          <span className="text-[10px] font-bold">
            Autori
          </span>
        </button>

        <button
          onClick={() => {
            setActiveTab(
              'settings'
            );

            setSelectedAuthor(
              null
            );
          }}
          className={`flex flex-col items-center gap-1 ${
            activeTab ===
            'settings'
              ? 'text-amber-800'
              : 'text-amber-900/40'
          }`}
        >
          <Settings className="w-5 h-5" />

          <span className="text-[10px] font-bold">
            Settings
          </span>
        </button>
      </nav>
    </div>
  );
}
