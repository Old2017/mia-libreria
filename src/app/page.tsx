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
} from 'lucide-react';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';

/* ============================================================
   TIPI
============================================================ */

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

/* ============================================================
   TIPI BARCODE DETECTOR
   Evita errori TypeScript nei progetti dove lib.dom
   non contiene BarcodeDetector.
============================================================ */

interface DetectedBarcode {
  rawValue?: string;
}

interface BarcodeDetectorInstance {
  detect(
    source: HTMLVideoElement
  ): Promise<DetectedBarcode[]>;
}

interface BarcodeDetectorConstructor {
  new (options?: {
    formats?: string[];
  }): BarcodeDetectorInstance;
}

declare global {
  interface Window {
    BarcodeDetector?: BarcodeDetectorConstructor;
  }
}

/* ============================================================
   COMPONENTE
============================================================ */

export default function LibraryApp() {
  /* ------------------------------------------------------------
     NAVIGAZIONE
  ------------------------------------------------------------ */

  const [activeTab, setActiveTab] = useState<
    'home' | 'read' | 'authors' | 'settings'
  >('home');

  const [books, setBooks] = useState<BookItem[]>([]);
  const [isLoaded, setIsLoaded] = useState(false);

  /* ------------------------------------------------------------
     MODALI
  ------------------------------------------------------------ */

  const [isAddModalOpen, setIsAddModalOpen] = useState(false);
  const [selectedBookDetail, setSelectedBookDetail] =
    useState<BookItem | null>(null);

  const [selectedAuthor, setSelectedAuthor] =
    useState<string | null>(null);

  const [homeSubView, setHomeSubView] = useState<
    'none' | 'classics' | 'genres'
  >('none');

  const [selectedGenreHome, setSelectedGenreHome] =
    useState<string | null>(null);

  /* ------------------------------------------------------------
     DRAG & DROP
  ------------------------------------------------------------ */

  const [draggedIndex, setDraggedIndex] =
    useState<number | null>(null);

  /* ------------------------------------------------------------
     ISBN MANUALE / API
  ------------------------------------------------------------ */

  const [isbnInput, setIsbnInput] = useState('');
  const [isSearchingIsbn, setIsSearchingIsbn] =
    useState(false);

  /* ------------------------------------------------------------
     CAMERA / SCANNER
  ------------------------------------------------------------ */

  const videoRef = useRef<HTMLVideoElement | null>(null);

  const [isCameraScannerOpen, setIsCameraScannerOpen] =
    useState(false);

  const [isCameraScanning, setIsCameraScanning] =
    useState(false);

  const [cameraError, setCameraError] = useState('');

  const [cameraStream, setCameraStream] =
    useState<MediaStream | null>(null);

  const cameraStreamRef = useRef<MediaStream | null>(null);

  const scannerRunningRef = useRef(false);

  /* ------------------------------------------------------------
     FILTRI
  ------------------------------------------------------------ */

  const [filterGenre, setFilterGenre] =
    useState<string>('all');

  const [filterFormat, setFilterFormat] =
    useState<string>('all');

  const [filterYear, setFilterYear] =
    useState<string>('all');

  const [searchQuery, setSearchQuery] =
    useState('');

  /* ------------------------------------------------------------
     ANNI
  ------------------------------------------------------------ */

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

  /* ------------------------------------------------------------
     FORM
  ------------------------------------------------------------ */

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

  /* ============================================================
     STORAGE
  ============================================================ */

  useEffect(() => {
    const saved = localStorage.getItem(
      'ios_library_books_v5'
    );

    if (saved) {
      try {
        setBooks(JSON.parse(saved));
      } catch (error) {
        console.error(
          'Errore caricamento storage',
          error
        );
      }
    }

    setIsLoaded(true);
  }, []);

  useEffect(() => {
    if (!isLoaded) return;

    localStorage.setItem(
      'ios_library_books_v5',
      JSON.stringify(books)
    );
  }, [books, isLoaded]);

  /* ============================================================
     STOP CAMERA
  ============================================================ */

  const stopCameraScanner = () => {
    scannerRunningRef.current = false;

    const stream = cameraStreamRef.current;

    if (stream) {
      stream.getTracks().forEach((track) => {
        track.stop();
      });
    }

    cameraStreamRef.current = null;
    setCameraStream(null);

    if (videoRef.current) {
      videoRef.current.pause();
      videoRef.current.srcObject = null;
    }

    setIsCameraScanning(false);
    setIsCameraScannerOpen(false);
  };

  /* ============================================================
     CLEANUP CAMERA
  ============================================================ */

  useEffect(() => {
    return () => {
      const stream = cameraStreamRef.current;

      if (stream) {
        stream.getTracks().forEach((track) => {
          track.stop();
        });
      }

      cameraStreamRef.current = null;
    };
  }, []);

  /* ============================================================
     AVVIO FOTOCAMERA
  ============================================================ */

  const startCameraScanner = async () => {
    setCameraError('');

    if (
      typeof navigator === 'undefined' ||
      !navigator.mediaDevices?.getUserMedia
    ) {
      setCameraError(
        'La fotocamera non è disponibile in questo browser.'
      );
      setIsCameraScannerOpen(true);
      return;
    }

    if (
      typeof window === 'undefined' ||
      !window.BarcodeDetector
    ) {
      setCameraError(
        'Questo browser non supporta la scansione automatica dei codici a barre. Inserisci l’ISBN manualmente.'
      );
      setIsCameraScannerOpen(true);
      return;
    }

    try {
      const stream =
        await navigator.mediaDevices.getUserMedia({
          video: {
            facingMode: {
              ideal: 'environment',
            },
            width: {
              ideal: 1280,
            },
            height: {
              ideal: 720,
            },
          },
          audio: false,
        });

      cameraStreamRef.current = stream;
      setCameraStream(stream);

      setIsCameraScannerOpen(true);
      setIsCameraScanning(true);

      scannerRunningRef.current = true;

      setTimeout(() => {
        if (!videoRef.current) return;

        videoRef.current.srcObject = stream;

        videoRef.current
          .play()
          .catch((error) => {
            console.error(
              'Errore avvio video:',
              error
            );
          });
      }, 100);
    } catch (error) {
      console.error(
        'Errore accesso fotocamera:',
        error
      );

      setCameraError(
        'Non è stato possibile accedere alla fotocamera. Controlla i permessi del browser.'
      );

      setIsCameraScannerOpen(true);
    }
  };

  /* ============================================================
     RICERCA GOOGLE BOOKS
  ============================================================ */

  const handleSearchBookByISBN = async (
    codeToSearch?: string
  ) => {
    const query = (
      codeToSearch || isbnInput
    )
      .replace(/[^0-9Xx]/g, '')
      .trim();

    if (!query) {
      alert('Inserisci o scansiona un codice ISBN.');
      return;
    }

    if (
      query.length !== 10 &&
      query.length !== 13
    ) {
      alert(
        'L’ISBN deve contenere 10 oppure 13 caratteri.'
      );
      return;
    }

    setIsSearchingIsbn(true);

    try {
      /*
       * Chiamata diretta a Google Books.
       *
       * Se preferisci passare da una API Route Next.js,
       * trovi la versione consigliata più sotto.
       */

      const res = await fetch(
        `https://www.googleapis.com/books/v1/volumes?q=isbn:${encodeURIComponent(
          query
        )}`
      );

      if (!res.ok) {
        throw new Error(
          `Google Books HTTP ${res.status}`
        );
      }

      const data = await res.json();

      if (
        !data.items ||
        data.items.length === 0
      ) {
        alert(
          'Nessun libro trovato per questo ISBN. Puoi inserire i dati manualmente.'
        );
        return;
      }

      const info =
        data.items[0]?.volumeInfo || {};

      const cover =
        info.imageLinks?.thumbnail ||
        info.imageLinks?.smallThumbnail ||
        '';

      setFormData((prev) => ({
        ...prev,

        title:
          info.title ||
          prev.title ||
          '',

        author: info.authors
          ? info.authors.join(', ')
          : prev.author || '',

        publisher:
          info.publisher ||
          prev.publisher ||
          '',

        publishYear:
          info.publishedDate
            ? info.publishedDate.substring(0, 4)
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
          cover.replace(
            'http:',
            'https:'
          ) ||
          prev.coverUrl ||
          '',

        isbn: query,
      }));

      setIsbnInput(query);
    } catch (error) {
      console.error(
        'Errore Google Books:',
        error
      );

      alert(
        'Errore durante la ricerca online. Puoi compilare i campi manualmente.'
      );
    } finally {
      setIsSearchingIsbn(false);
    }
  };

  /* ============================================================
     SCANSIONE BARCODE
  ============================================================ */

  useEffect(() => {
    if (
      !isCameraScanning ||
      !isCameraScannerOpen ||
      !videoRef.current ||
      !window.BarcodeDetector
    ) {
      return;
    }

    const detector =
      new window.BarcodeDetector({
        formats: [
          'ean_13',
          'ean_8',
        ],
      });

    let animationFrame = 0;
    let cancelled = false;

    const scan = async () => {
      if (
        cancelled ||
        !scannerRunningRef.current ||
        !videoRef.current
      ) {
        return;
      }

      try {
        const barcodes =
          await detector.detect(
            videoRef.current
          );

        if (barcodes.length > 0) {
          const rawValue =
            barcodes[0]?.rawValue || '';

          const isbn = rawValue.replace(
            /[^0-9Xx]/g,
            ''
          );

          if (
            isbn.length === 10 ||
            isbn.length === 13
          ) {
            cancelled = true;
            scannerRunningRef.current = false;

            stopCameraScanner();

            setIsbnInput(isbn);

            await handleSearchBookByISBN(
              isbn
            );

            return;
          }
        }
      } catch (error) {
        console.error(
          'Errore durante scansione:',
          error
        );
      }

      if (
        !cancelled &&
        scannerRunningRef.current
      ) {
        animationFrame =
          requestAnimationFrame(scan);
      }
    };

    scan();

    return () => {
      cancelled = true;

      if (animationFrame) {
        cancelAnimationFrame(
          animationFrame
        );
      }
    };
  }, [
    isCameraScanning,
    isCameraScannerOpen,
  ]);

  /* ============================================================
     FORM RESET
  ============================================================ */

  const resetForm = () => {
    stopCameraScanner();

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
    setCameraError('');
  };

  /* ============================================================
     SALVA / MODIFICA LIBRO
  ============================================================ */

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

      genre: formData.genre || '',

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
        formData.isbn ||
        isbnInput ||
        '',

      createdAt:
        formData.createdAt ||
        Date.now(),
    };

    if (formData.id) {
      setBooks((prev) =>
        prev.map((book) =>
          book.id === formData.id
            ? newBook
            : book
        )
      );
    } else {
      setBooks((prev) => [
        newBook,
        ...prev,
      ]);
    }

    setIsAddModalOpen(false);
    setSelectedBookDetail(null);

    resetForm();
  };

  /* ============================================================
     MODIFICA
  ============================================================ */

  const handleEditBook = (
    book: BookItem
  ) => {
    setFormData(book);
    setIsbnInput(book.isbn || '');
    setSelectedBookDetail(null);
    setIsAddModalOpen(true);
  };

  /* ============================================================
     ELIMINA
  ============================================================ */

  const handleDeleteBook = (
    id: string
  ) => {
    if (
      confirm(
        'Sei sicuro di voler eliminare questo libro?'
      )
    ) {
      setBooks((prev) =>
        prev.filter(
          (book) => book.id !== id
        )
      );

      if (
        selectedBookDetail?.id === id
      ) {
        setSelectedBookDetail(null);
      }
    }
  };

  /* ============================================================
     DRAG & DROP
  ============================================================ */

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
        (book) => book.isRead
      );

    const itemToMove =
      readOnlyBooks[draggedIndex];

    if (!itemToMove) return;

    const updated = [...books];

    const sourceGlobalIdx =
      updated.findIndex(
        (book) =>
          book.id === itemToMove.id
      );

    const targetBook =
      readOnlyBooks[index];

    if (!targetBook) return;

    const targetGlobalIdx =
      updated.findIndex(
        (book) =>
          book.id === targetBook.id
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

  /* ============================================================
     EXPORT EXCEL
  ============================================================ */

  const exportToExcel = () => {
    const dataToExport =
      books.map((book) => ({
        Titolo: book.title,
        Autore: book.author,
        'Paese di Pubblicazione':
          book.publishCountry || '-',
        Classico: book.isClassic
          ? 'Sì'
          : 'No',
        Stato: book.isRead
          ? 'Letto'
          : 'In Biblioteca',
        Formato:
          book.format ===
          'cartaceo'
            ? 'Cartaceo'
            : book.format ===
                'ebook'
              ? 'eBook'
              : 'eBook + Cartaceo',
        Editore:
          book.publisher,
        'Anno Pubblicazione':
          book.publishYear,
        Genere:
          book.genre,
        'Serie / Tag':
          book.seriesTag,
        Volume:
          book.volume || '-',
        Pagine:
          book.pages,
        'Mese e Anno di Lettura':
          book.readMonthYear,
        ISBN:
          book.isbn,
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

  /* ============================================================
     EXPORT PDF
  ============================================================ */

  const exportToPDF = () => {
    const doc = new jsPDF();

    doc.text(
      'La Mia Biblioteca - Report',
      14,
      15
    );

    const tableData =
      books.map((book) => [
        book.title,
        book.author,
        book.publishCountry ||
          '-',
        book.isClassic
          ? 'Sì'
          : 'No',
        book.isRead
          ? 'Letto'
          : 'In Libreria',
        book.format ===
        'cartaceo'
          ? 'Cartaceo'
          : book.format ===
              'ebook'
            ? 'eBook'
            : 'eBook + Cartaceo',
        book.genre || '-',
        book.readMonthYear ||
          '-',
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

  /* ============================================================
     BACKUP
  ============================================================ */

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

  /* ============================================================
     IMPORT BACKUP
  ============================================================ */

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
              event.target
                ?.result as string
            );

          if (
            Array.isArray(parsed)
          ) {
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

  /* ============================================================
     CANCELLA TUTTO
  ============================================================ */

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

  /* ============================================================
     STATISTICHE
  ============================================================ */

  const totalBooks = books.length;

  const totalCartacei =
    books.filter(
      (book) =>
        book.format ===
          'cartaceo' ||
        book.format ===
          'ebook_and_paper'
    ).length;

  const totalEbook =
    books.filter(
      (book) =>
        book.format ===
          'ebook' ||
        book.format ===
          'ebook_and_paper'
    ).length;

  const readBooks =
    books.filter(
      (book) => book.isRead
    );

  const readBooksCount =
    readBooks.length;

  const readCartacei =
    readBooks.filter(
      (book) =>
        book.format ===
          'cartaceo' ||
        book.format ===
          'ebook_and_paper'
    ).length;

  const readEbook =
    readBooks.filter(
      (book) =>
        book.format ===
          'ebook' ||
        book.format ===
          'ebook_and_paper'
    ).length;

  const readThisYearBooks =
    books.filter((book) => {
      if (!book.isRead)
        return false;

      if (book.readYear) {
        return (
          book.readYear ===
          currentYearNum
        );
      }

      if (
        book.readMonthYear &&
        book.readMonthYear.includes(
          currentYearNum.toString()
        )
      ) {
        return true;
      }

      return false;
    });

  const readThisYearCount =
    readThisYearBooks.length;

  const readThisYearCartacei =
    readThisYearBooks.filter(
      (book) =>
        book.format ===
          'cartaceo' ||
        book.format ===
          'ebook_and_paper'
    ).length;

  const readThisYearEbook =
    readThisYearBooks.filter(
      (book) =>
        book.format ===
          'ebook' ||
        book.format ===
          'ebook_and_paper'
    ).length;

  /* ============================================================
     FILTRI LIBRI LETTI
  ============================================================ */

  const readBooksFiltered =
    books
      .filter(
        (book) => book.isRead
      )
      .filter((book) =>
        filterGenre === 'all'
          ? true
          : book.genre ===
            filterGenre
      )
      .filter((book) => {
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
            book.format ===
            'cartaceo'
          );
        }

        if (
          filterFormat ===
          'ebook'
        ) {
          return (
            book.format ===
              'ebook' ||
            book.format ===
              'ebook_and_paper'
          );
        }

        return true;
      })
      .filter((book) => {
        if (
          filterYear === 'all'
        ) {
          return true;
        }

        return (
          book.readYear?.toString() ===
            filterYear ||
          book.readMonthYear?.includes(
            filterYear
          )
        );
      })
      .filter((book) =>
        searchQuery === ''
          ? true
          : book.title
              .toLowerCase()
              .includes(
                searchQuery.toLowerCase()
              ) ||
            book.author
              .toLowerCase()
              .includes(
                searchQuery.toLowerCase()
              )
      );

  /* ============================================================
     AUTORI
  ============================================================ */

  const authorsMap =
    books.reduce(
      (
        acc,
        book
      ) => {
        const authorName =
          book.author.trim() ||
          'Autore Sconosciuto';

        if (!acc[authorName]) {
          acc[authorName] = [];
        }

        acc[authorName].push(
          book
        );

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

  /* ============================================================
     GENERI
  ============================================================ */

  const genresMap =
    books.reduce(
      (
        acc,
        book
      ) => {
        const genreName =
          book.genre?.trim() ||
          'Generico / Altro';

        if (!acc[genreName]) {
          acc[genreName] = [];
        }

        acc[genreName].push(
          book
        );

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
          .map(
            (book) =>
              book.genre
          )
          .filter(Boolean)
      )
    );

  const availableYears =
    Array.from(
      new Set(
        books
          .map((book) => {
            if (book.readYear) {
              return book.readYear.toString();
            }

            const match =
              book.readMonthYear?.match(
                /\d{4}/
              );

            return match
              ? match[0]
              : null;
          })
          .filter(Boolean)
      )
    );

  /* ============================================================
     FORMAT LABEL
  ============================================================ */

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

  /* ============================================================
     RENDER GRUPPI AUTORE
  ============================================================ */

  const renderAuthorGroup = (
    bookList: BookItem[]
  ) => {
    const map =
      bookList.reduce(
        (
          acc,
          book
        ) => {
          const author =
            book.author.trim() ||
            'Autore Sconosciuto';

          if (!acc[author]) {
            acc[author] = [];
          }

          acc[author].push(
            book
          );

          return acc;
        },
        {} as Record<
          string,
          BookItem[]
        >
      );

    return Object.keys(
      map
    )
      .sort()
      .map(
        (author) => (
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
                  (
                    a.volume || ''
                  ).localeCompare(
                    b.volume || ''
                  )
                )
                .map(
                  (book) => (
                    <div
                      key={
                        book.id
                      }
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

                      <p className="text-[10px] text-amber-800/60 mt-0.5">
                        {book.publishYear
                          ? `Anno: ${book.publishYear}`
                          : ''}
                      </p>
                    </div>
                  )
                )}
            </div>
          </div>
        )
      );
  };

  /* ============================================================
     RENDER
  ============================================================ */

  return (
    <div className="min-h-screen bg-[#FBF9F5] text-amber-950 font-sans pb-24 select-none">

      {/* ======================================================
          HEADER
      ====================================================== */}

      <header className="sticky top-0 z-20 bg-[#FBF9F5]/90 backdrop-blur-md border-b border-amber-900/10 px-5 py-3.5 flex justify-between items-center">
        <div>
          <span className="text-[11px] font-bold text-amber-800/70 uppercase tracking-widest block">
            {activeTab ===
              'home' &&
              'La Mia Collezione'}

            {activeTab ===
              'read' &&
              'Cronologia Letture'}

            {activeTab ===
              'authors' &&
              'Catalogo Autori'}

            {activeTab ===
              'settings' &&
              'Gestione Dati'}
          </span>

          <h1 className="text-2xl font-serif font-extrabold tracking-tight text-amber-950">
            {activeTab ===
              'home' &&
              'Home'}

            {activeTab ===
              'read' &&
              'Libri Letti'}

            {activeTab ===
              'authors' &&
              (selectedAuthor
                ? selectedAuthor
                : 'Autori')}

            {activeTab ===
              'settings' &&
              'Impostazioni'}
          </h1>
        </div>

        {activeTab ===
          'read' && (
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

      {/* ======================================================
          HOME
      ====================================================== */}

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
                      (book) =>
                        book.isClassic
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
                        Generi nella Biblioteca
                      </h2>

                      {sortedGenres.map(
                        (genre) => (
                          <div
                            key={genre}
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
                        ← Tutti i generi
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
                      {
                        currentYearNum
                      }
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
                      {
                        totalBooks
                      }
                    </span>

                    <span className="text-xs font-bold text-amber-800/60 uppercase tracking-wider block mt-1">
                      In Biblioteca
                    </span>

                    <div className="mt-2 pt-2 border-t border-amber-900/10 text-xs font-semibold text-amber-900/80">
                      {
                        totalCartacei
                      }{' '}
                      Cartacei e{' '}
                      {
                        totalEbook
                      }{' '}
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
                      {
                        readEbook
                      }{' '}
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
                            (book) =>
                              book.isClassic
                          ).length
                        }{' '}
                        libri conservati
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

      {/* ======================================================
          LIBRI LETTI
      ====================================================== */}

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
                value={filterGenre}
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
                  (genre) => (
                    <option
                      key={genre}
                      value={genre}
                    >
                      {genre}
                    </option>
                  )
                )}
              </select>

              <select
                value={filterFormat}
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
                  (year) => (
                    <option
                      key={year}
                      value={year!}
                    >
                      {year}
                    </option>
                  )
                )}
              </select>

            </div>
          </div>

          <p className="text-[11px] text-amber-800/50 font-medium px-1">
            Trascina tramite l'icona a sinistra per riordinare l'elenco dei libri letti.
          </p>

          <div className="space-y-3">

            {readBooksFiltered.length ===
            0 ? (
              <div className="text-center py-12 bg-[#FFFDF9] rounded-3xl border border-dashed border-amber-900/20">
                <BookOpen className="w-8 h-8 text-amber-800/30 mx-auto mb-2" />

                <p className="text-xs text-amber-800/60 font-medium">
                  Nessun libro letto trovato.
                </p>
              </div>
            ) : (
              readBooksFiltered.map(
                (
                  book,
                  index
                ) => (
                  <div
                    key={book.id}
                    draggable
                    onDragStart={() =>
                      handleDragStart(
                        index
                      )
                    }
                    onDragOver={(e) =>
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

      {/* ======================================================
          AUTORI
      ====================================================== */}

      {activeTab ===
        'authors' && (
        <div className="p-4 space-y-4 max-w-lg mx-auto">

          {!selectedAuthor ? (
            <div className="space-y-2">

              {sortedAuthors.length ===
              0 ? (
                <div className="text-center py-12 bg-[#FFFDF9] rounded-3xl border border-dashed border-amber-900/20">
                  <Users className="w-8 h-8 text-amber-800/30 mx-auto mb-2" />

                  <p className="text-xs text-amber-800/60 font-medium">
                    Nessun autore presente.
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
                ← Torna alla lista autori
              </button>

              <div className="grid grid-cols-2 gap-3.5">

                {authorsMap[
                  selectedAuthor
                ]?.map((book) => (
                  <div
                    key={book.id}
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

      {/* ======================================================
          SETTINGS
      ====================================================== */}

      {activeTab ===
        'settings' && (
        <div className="p-4 space-y-5 max-w-lg mx-auto">

          <div className="bg-[#FFFDF9] rounded-3xl p-4 shadow-sm border border-amber-900/10 space-y-3">

            <h2 className="text-xs font-bold text-amber-800/60 uppercase tracking-wider">
              Esporta la tua Biblioteca
            </h2>

            <button
              onClick={
                exportToExcel
              }
              className="w-full p-3 bg-emerald-50 text-emerald-900 rounded-2xl font-bold text-xs flex items-center gap-3 border border-emerald-200"
            >
              <FileSpreadsheet className="w-5 h-5 text-emerald-700" />
              Esporta in Foglio Excel (.xlsx)
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
              Backup & Ripristino Dati
            </h2>

            <button
              onClick={
                exportBackup
              }
              className="w-full p-3 bg-amber-100/60 text-amber-950 rounded-2xl font-bold text-xs flex items-center gap-3 border border-amber-200"
            >
              <Download className="w-5 h-5 text-amber-800" />
              Salva Backup Dati (JSON)
            </button>

            <label className="w-full p-3 bg-stone-100/80 text-stone-900 rounded-2xl font-bold text-xs flex items-center gap-3 border border-stone-200 cursor-pointer">
              <Upload className="w-5 h-5 text-stone-700" />
              Ripristina Backup da File

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
              Cancella Intera Biblioteca
            </button>

          </div>
        </div>
      )}

      {/* ======================================================
          MODALE DETTAGLIO LIBRO
      ====================================================== */}

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
                  Paese di Pubblicazione:
                </span>

                <span className="font-bold">
                  {
                    selectedBookDetail.publishCountry ||
                    '-'
                  }
                </span>
              </div>

              <div>
                <span className="block text-amber-800/50 font-semibold text-[10px]">
                  Editore:
                </span>

                <span className="font-bold">
                  {
                    selectedBookDetail.publisher ||
                    '-'
                  }
                </span>
              </div>

              <div>
                <span className="block text-amber-800/50 font-semibold text-[10px]">
                  Anno Pubblicazione:
                </span>

                <span className="font-bold">
                  {
                    selectedBookDetail.publishYear ||
                    '-'
                  }
                </span>
              </div>

              <div>
                <span className="block text-amber-800/50 font-semibold text-[10px]">
                  Pagine:
                </span>

                <span className="font-bold">
                  {
                    selectedBookDetail.pages ||
                    '-'
                  }
                </span>
              </div>

              <div>
                <span className="block text-amber-800/50 font-semibold text-[10px]">
                  Genere:
                </span>

                <span className="font-bold">
                  {
                    selectedBookDetail.genre ||
                    '-'
                  }
                </span>
              </div>

              <div>
                <span className="block text-amber-800/50 font-semibold text-[10px]">
                  Serie / Tag:
                </span>

                <span className="font-bold">
                  {
                    selectedBookDetail.seriesTag ||
                    '-'
                  }
                </span>
              </div>

              <div>
                <span className="block text-amber-800/50 font-semibold text-[10px]">
                  Volume:
                </span>

                <span className="font-bold">
                  {
                    selectedBookDetail.volume ||
                    '-'
                  }
                </span>
              </div>

              {selectedBookDetail.isbn && (
                <div>
                  <span className="block text-amber-800/50 font-semibold text-[10px]">
                    ISBN:
                  </span>

                  <span className="font-bold">
                    {
                      selectedBookDetail.isbn
                    }
                  </span>
                </div>
              )}

              {selectedBookDetail.isRead && (
                <div className="col-span-2 border-t border-amber-900/10 pt-2">
                  <span className="block text-amber-800/50 font-semibold text-[10px]">
                    Mese & Anno di Lettura:
                  </span>

                  <span className="font-bold">
                    {
                      selectedBookDetail.readMonthYear ||
                      '-'
                    }
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

      {/* ======================================================
          MODALE CAMERA ISBN
      ====================================================== */}

      {isCameraScannerOpen && (
        <div className="fixed inset-0 z-[70] bg-black flex flex-col">

          <div className="flex items-center justify-between p-4 bg-black text-white">

            <div>
              <h2 className="font-serif font-bold text-base">
                Scansiona ISBN
              </h2>

              <p className="text-[10px] text-white/60 mt-0.5">
                Inquadra il codice a barre sul retro del libro
              </p>
            </div>

            <button
              type="button"
              onClick={
                stopCameraScanner
              }
              className="p-2 bg-white/10 rounded-full"
            >
              <X className="w-5 h-5" />
            </button>

          </div>

          <div className="flex-1 relative flex items-center justify-center overflow-hidden">

            {!cameraError && (
              <video
                ref={videoRef}
                muted
                playsInline
                autoPlay
                className="absolute inset-0 w-full h-full object-cover"
              />
            )}

            {!cameraError && (
              <>
                <div className="relative z-10 w-[82%] max-w-sm aspect-[2/1] border-2 border-white rounded-2xl shadow-[0_0_0_9999px_rgba(0,0,0,0.45)]">

                  <div className="absolute left-0 right-0 top-1/2 h-0.5 bg-red-500 shadow-lg" />

                  <div className="absolute -top-1 -left-1 w-8 h-8 border-t-4 border-l-4 border-amber-400 rounded-tl-xl" />

                  <div className="absolute -top-1 -right-1 w-8 h-8 border-t-4 border-r-4 border-amber-400 rounded-tr-xl" />

                  <div className="absolute -bottom-1 -left-1 w-8 h-8 border-b-4 border-l-4 border-amber-400 rounded-bl-xl" />

                  <div className="absolute -bottom-1 -right-1 w-8 h-8 border-b-4 border-r-4 border-amber-400 rounded-br-xl" />

                </div>

                <div className="absolute bottom-8 left-0 right-0 z-20 text-center">

                  <span className="inline-block px-4 py-2 bg-black/60 backdrop-blur-md rounded-full text-white text-xs font-medium">
                    Posiziona il codice ISBN all'interno del riquadro
                  </span>

                </div>
              </>
            )}

            {cameraError && (
              <div className="relative z-20 max-w-sm mx-auto px-6 text-center">

                <Scan className="w-12 h-12 text-amber-400 mx-auto mb-4" />

                <p className="text-white text-sm font-medium">
                  {cameraError}
                </p>

                <button
                  type="button"
                  onClick={
                    stopCameraScanner
                  }
                  className="mt-5 px-5 py-3 bg-white text-black rounded-xl text-xs font-bold"
                >
                  Torna al form
                </button>

              </div>
            )}

          </div>

          {!cameraError && (
            <div className="p-4 bg-black text-center">

              <button
                type="button"
                onClick={
                  stopCameraScanner
                }
                className="px-6 py-3 bg-white/10 text-white rounded-xl text-xs font-bold"
              >
                Inserisci ISBN manualmente
              </button>

            </div>
          )}

        </div>
      )}

      {/* ======================================================
          MODALE AGGIUNTA / MODIFICA
      ====================================================== */}

      {isAddModalOpen && (
        <div className="fixed inset-0 z-50 bg-amber-950/40 backdrop-blur-sm flex items-end sm:items-center justify-center p-0 sm:p-4">

          <div className="bg-[#FFFDF9] w-full max-w-lg rounded-t-3xl sm:rounded-3xl max-h-[90vh] overflow-y-auto p-6 space-y-4 shadow-2xl relative border border-amber-900/10">

            <div className="flex justify-between items-center border-b border-amber-900/10 pb-3">

              <h2 className="text-base font-serif font-bold text-amber-950">
                {formData.id
                  ? 'Modifica Libro'
                  : 'Aggiungi Nuovo Libro'}
              </h2>

              <button
                onClick={() => {
                  stopCameraScanner();
                  setIsAddModalOpen(
                    false
                  );
                }}
                className="p-2 text-amber-800/40"
              >
                <X className="w-5 h-5" />
              </button>

            </div>

            {/* =================================================
                ISBN / CAMERA
            ================================================= */}

            <div className="bg-amber-100/50 p-3.5 rounded-2xl space-y-3 border border-amber-900/10">

              <div className="flex justify-between items-center">

                <div>
                  <span className="text-[10px] font-bold text-amber-900 uppercase tracking-wider block">
                    Inserimento automatico
                  </span>

                  <span className="text-[10px] text-amber-800/60 font-medium">
                    Scansiona il codice ISBN oppure inseriscilo manualmente
                  </span>
                </div>

                <Scan className="w-5 h-5 text-amber-800" />

              </div>

              <div className="grid grid-cols-2 gap-2">

                <button
                  type="button"
                  onClick={
                    startCameraScanner
                  }
                  className="py-2.5 bg-amber-800 text-amber-50 rounded-xl text-xs font-bold flex items-center justify-center gap-2 active:scale-95 transition-transform"
                >
                  <Scan className="w-4 h-4" />
                  Scansiona ISBN
                </button>

                <button
                  type="button"
                  onClick={() =>
                    handleSearchBookByISBN()
                  }
                  disabled={
                    isSearchingIsbn ||
                    !isbnInput.trim()
                  }
                  className="py-2.5 bg-[#FFFDF9] text-amber-900 border border-amber-900/10 rounded-xl text-xs font-bold flex items-center justify-center gap-2 disabled:opacity-40"
                >
                  <Search className="w-4 h-4" />

                  {isSearchingIsbn
                    ? 'Ricerca...'
                    : 'Cerca ISBN'}
                </button>

              </div>

              <input
                type="text"
                inputMode="numeric"
                autoComplete="off"
                placeholder="ISBN 10 o ISBN 13"
                value={isbnInput}
                onChange={(e) =>
                  setIsbnInput(
                    e.target.value.replace(
                      /[^0-9Xx-]/g,
                      ''
                    )
                  )
                }
                onKeyDown={(e) => {
                  if (
                    e.key === 'Enter'
                  ) {
                    e.preventDefault();

                    handleSearchBookByISBN();
                  }
                }}
                className="w-full p-2.5 bg-[#FFFDF9] border border-amber-900/10 rounded-xl text-xs font-medium focus:outline-none focus:ring-2 focus:ring-amber-800/20"
              />

              {formData.isbn && (
                <div className="flex items-center gap-2 text-[10px] text-emerald-700 font-semibold">
                  <CheckCircle2 className="w-3.5 h-3.5" />

                  ISBN associato:
                  {' '}
                  {
                    formData.isbn
                  }
                </div>
              )}

            </div>

            {/* =================================================
                FORM
            ================================================= */}

            <form
              onSubmit={
                handleSaveBook
              }
              className="space-y-3 text-xs"
            >

              <input
                type="text"
                placeholder="Titolo *"
                value={
                  formData.title || ''
                }
                onChange={(e) =>
                  setFormData({
                    ...formData,
                    title:
                      e.target.value,
                  })
                }
                className="w-full p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl font-medium focus:outline-none"
                required
              />

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
                        e.target.value,
                    })
                  }
                  className="p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl font-medium focus:outline-none"
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
                        e.target.value,
                    })
                  }
                  className="p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl font-medium focus:outline-none"
                />

              </div>

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
                      e.target.value,
                  })
                }
                className="w-full p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl font-medium focus:outline-none"
              />

              {formData.coverUrl && (
                <div className="flex items-center gap-3 p-2 bg-amber-50 rounded-xl border border-amber-900/10">

                  <img
                    src={
                      formData.coverUrl
                    }
                    alt="Anteprima copertina"
                    className="w-12 h-16 object-cover rounded-lg border border-amber-900/10"
                    onError={(e) => {
                      e.currentTarget.style.display =
                        'none';
                    }}
                  />

                  <div>
                    <p className="text-[10px] font-bold text-amber-900">
                      Anteprima copertina
                    </p>

                    <p className="text-[9px] text-amber-800/60">
                      Recuperata automaticamente dall'ISBN
                    </p>
                  </div>

                </div>
              )}

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
                        e.target.value,
                    })
                  }
                  className="p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl font-medium focus:outline-none"
                />

                <input
                  type="text"
                  placeholder="Anno Pubblicazione"
                  value={
                    formData.publishYear ||
                    ''
                  }
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      publishYear:
                        e.target.value,
                    })
                  }
                  className="p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl font-medium focus:outline-none"
                />

              </div>

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
                        e.target.value,
                    })
                  }
                  className="p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl font-medium focus:outline-none"
                />

                <input
                  type="text"
                  inputMode="numeric"
                  pattern="[0-9]*"
                  placeholder="Numero Pagine"
                  value={
                    formData.pages || ''
                  }
                  onChange={(e) => {
                    const value =
                      e.target.value.replace(
                        /\D/g,
                        ''
                      );

                    setFormData({
                      ...formData,
                      pages: value
                        ? parseInt(
                            value,
                            10
                          )
                        : undefined,
                    });
                  }}
                  className="p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl font-medium focus:outline-none"
                />

              </div>

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
                        e.target.value,
                    })
                  }
                  className="p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl font-medium focus:outline-none"
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
                        e.target.value,
                    })
                  }
                  className="p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl font-medium focus:outline-none"
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
                        e.target.checked,
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
                      Acquistato anche in formato Cartaceo
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
                        e.target.checked,
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
                        onChange={(e) =>
                          setFormData({
                            ...formData,
                            readMonth:
                              e.target.value,
                          })
                        }
                        className="w-full p-2.5 bg-[#FFFDF9] border border-amber-900/10 rounded-xl font-semibold text-amber-950 focus:outline-none"
                      >
                        {MONTHS.map(
                          (
                            month
                          ) => (
                            <option
                              key={
                                month
                              }
                              value={
                                month
                              }
                            >
                              {
                                month
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
                        onChange={(e) =>
                          setFormData({
                            ...formData,
                            readYear:
                              parseInt(
                                e.target.value,
                                10
                              ),
                          })
                        }
                        className="w-full p-2.5 bg-[#FFFDF9] border border-amber-900/10 rounded-xl font-semibold text-amber-950 focus:outline-none"
                      >
                        {yearsList.map(
                          (
                            year
                          ) => (
                            <option
                              key={
                                year
                              }
                              value={
                                year
                              }
                            >
                              {
                                year
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

      {/* ======================================================
          TAB BAR
      ====================================================== */}

      <nav className="fixed bottom-0 left-0 right-0 z-30 bg-[#FBF9F5]/90 backdrop-blur-md border-t border-amber-900/10 flex justify-around py-2.5 max-w-lg mx-auto">

        <button
          onClick={() => {
            setActiveTab('home');
            setSelectedAuthor(null);
            setHomeSubView(
              'none'
            );
          }}
          className={`flex flex-col items-center gap-1 ${
            activeTab ===
            'home'
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
            setActiveTab('read');
            setSelectedAuthor(null);
          }}
          className={`flex flex-col items-center gap-1 ${
            activeTab ===
            'read'
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
            setSelectedAuthor(null);
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
            setSelectedAuthor(null);
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
