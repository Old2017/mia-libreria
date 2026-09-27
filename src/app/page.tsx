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
  Camera,
  Trash2,
  Edit,
  Download,
  Upload,
  CheckCircle2,
  Book,
  Tablet,
  ChevronRight,
  Star,
  X,
  FileSpreadsheet,
  FileText,
  Calendar,
  GripVertical,
  Layers,
  Sparkles
} from 'lucide-react';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import { BrowserMultiFormatReader } from '@zxing/library';

// --- INTERFACCIA DATI LIBRO ---
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
  'Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno',
  'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'
];

export default function LibraryApp() {
  const [activeTab, setActiveTab] = useState<'home' | 'read' | 'authors' | 'settings'>('home');
  const [books, setBooks] = useState<BookItem[]>([]);
  const [isLoaded, setIsLoaded] = useState(false);

  // Modali e Viste Interne Home
  const [isAddModalOpen, setIsAddModalOpen] = useState(false);
  const [selectedBookDetail, setSelectedBookDetail] = useState<BookItem | null>(null);
  const [selectedAuthor, setSelectedAuthor] = useState<string | null>(null);
  const [homeSubView, setHomeSubView] = useState<'none' | 'classics' | 'genres'>('none');
  const [selectedGenreHome, setSelectedGenreHome] = useState<string | null>(null);

  // Scanner Fotocamera
  const [isScanningCamera, setIsScanningCamera] = useState(false);
  const videoRef = useRef<HTMLVideoElement | null>(null);
  const codeReaderRef = useRef<BrowserMultiFormatReader | null>(null);

  // Drag & Drop State (Sezione Letti)
  const [draggedIndex, setDraggedIndex] = useState<number | null>(null);

  // Scanner ISBN
  const [isbnInput, setIsbnInput] = useState('');
  const [isSearchingIsbn, setIsSearchingIsbn] = useState(false);

  // Filtri Sezione Letti
  const [filterGenre, setFilterGenre] = useState<string>('all');
  const [filterFormat, setFilterFormat] = useState<string>('all');
  const [filterYear, setFilterYear] = useState<string>('all');
  const [searchQuery, setSearchQuery] = useState('');

  // Anni dal 2023 fino all'anno corrente
  const currentYearNum = new Date().getFullYear();
  const startYear = 2023;
  const yearsList = Array.from(
    { length: Math.max(1, currentYearNum - startYear + 1) },
    (_, i) => currentYearNum - i
  );

  // Form State
  const [formData, setFormData] = useState<Partial<BookItem>>({
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
  });

  // Caricamento Dati
  useEffect(() => {
    const saved = localStorage.getItem('ios_library_books_v5');
    if (saved) {
      try {
        setBooks(JSON.parse(saved));
      } catch (e) {
        console.error('Errore caricamento storage', e);
      }
    }
    setIsLoaded(true);
  }, []);

  // Salvataggio Dati
  useEffect(() => {
    if (isLoaded) {
      localStorage.setItem('ios_library_books_v5', JSON.stringify(books));
    }
  }, [books, isLoaded]);

  // Gestione spegnimento fotocamera alla chiusura modale
  useEffect(() => {
    if (!isAddModalOpen && isScanningCamera) {
      stopCameraScan();
    }
  }, [isAddModalOpen]);

  const startCameraScan = async () => {
    setIsScanningCamera(true);
    codeReaderRef.current = new BrowserMultiFormatReader();
    try {
      const videoInputDevices = await codeReaderRef.current.listVideoInputDevices();
      const selectedDeviceId = videoInputDevices.length > 0 ? videoInputDevices[0].deviceId : undefined;
      
      if (videoRef.current) {
        codeReaderRef.current.decodeFromVideoDevice(
          selectedDeviceId,
          videoRef.current,
          (result) => {
            if (result) {
              const scannedIsbn = result.getText();
              setIsbnInput(scannedIsbn);
              stopCameraScan();
              handleSearchBookByISBN(scannedIsbn);
            }
          }
        );
      }
    } catch (err) {
      console.error(err);
      alert('Impossibile accedere alla fotocamera. Controlla i permessi del browser.');
      setIsScanningCamera(false);
    }
  };

  const stopCameraScan = () => {
    if (codeReaderRef.current) {
      codeReaderRef.current.reset();
    }
    setIsScanningCamera(false);
  };

  // Statistiche Totali Biblioteca
  const totalBooks = books.length;
  const totalCartacei = books.filter((b) => b.format === 'cartaceo' || b.format === 'ebook_and_paper').length;
  const totalEbook = books.filter((b) => b.format === 'ebook' || b.format === 'ebook_and_paper').length;

  // Statistiche Libri Letti
  const readBooks = books.filter((b) => b.isRead);
  const readBooksCount = readBooks.length;
  const readCartacei = readBooks.filter((b) => b.format === 'cartaceo' || b.format === 'ebook_and_paper').length;
  const readEbook = readBooks.filter((b) => b.format === 'ebook' || b.format === 'ebook_and_paper').length;

  // Statistiche Anno in Corso
  const readThisYearBooks = books.filter((b) => {
    if (!b.isRead) return false;
    if (b.readYear) return b.readYear === currentYearNum;
    if (b.readMonthYear && b.readMonthYear.includes(currentYearNum.toString())) return true;
    return false;
  });
  const readThisYearCount = readThisYearBooks.length;
  const readThisYearCartacei = readThisYearBooks.filter((b) => b.format === 'cartaceo' || b.format === 'ebook_and_paper').length;
  const readThisYearEbook = readThisYearBooks.filter((b) => b.format === 'ebook' || b.format === 'ebook_and_paper').length;

  // Ricerca Google Books per ISBN
  const handleSearchBookByISBN = async (codeToSearch?: string) => {
    const query = codeToSearch || isbnInput;
    if (!query) return;
    setIsSearchingIsbn(true);

    try {
      const res = await fetch(`https://www.googleapis.com/books/v1/volumes?q=isbn:${query}`);
      const data = await res.json();

      if (data.items && data.items.length > 0) {
        const info = data.items[0].volumeInfo;
        setFormData((prev) => ({
          ...prev,
          title: info.title || prev.title || '',
          author: info.authors ? info.authors.join(', ') : prev.author || '',
          publisher: info.publisher || prev.publisher || '',
          publishYear: info.publishedDate ? info.publishedDate.substring(0, 4) : prev.publishYear || '',
          pages: info.pageCount || prev.pages || undefined,
          genre: info.categories ? info.categories[0] : prev.genre || '',
          coverUrl: info.imageLinks?.thumbnail?.replace('http:', 'https:') || prev.coverUrl || '',
          isbn: query,
        }));
      } else {
        alert('Nessun dato trovato sul web per questo codice. Puoi inserire i dettagli manualmente.');
      }
    } catch (err) {
      console.error(err);
      alert('Errore durante la ricerca online. Compila i campi manualmente.');
    } finally {
      setIsSearchingIsbn(false);
    }
  };

  // Salva o Aggiorna Libro
  const handleSaveBook = (e: React.FormEvent) => {
    e.preventDefault();
    if (!formData.title || !formData.author) {
      alert('Inserisci almeno Titolo e Autore.');
      return;
    }

    const formattedReadMonthYear = formData.isRead && formData.readMonth && formData.readYear 
      ? `${formData.readMonth} ${formData.readYear}` 
      : '';

    const newBook: BookItem = {
      id: formData.id || Date.now().toString(),
      title: formData.title || '',
      author: formData.author || '',
      publishCountry: formData.publishCountry || '',
      coverUrl: formData.coverUrl || '',
      publisher: formData.publisher || '',
      publishYear: formData.publishYear || '',
      pages: Number(formData.pages) || undefined,
      genre: formData.genre || '',
      seriesTag: formData.seriesTag || '',
      volume: formData.volume || '',
      isClassic: !!formData.isClassic,
      format: formData.format || 'cartaceo',
      isRead: !!formData.isRead,
      readMonth: formData.readMonth || 'Gennaio',
      readYear: formData.readYear || currentYearNum,
      readMonthYear: formattedReadMonthYear,
      rating: Number(formData.rating) || 5,
      notes: formData.notes || '',
      isbn: formData.isbn || '',
      createdAt: formData.createdAt || Date.now(),
    };

    if (formData.id) {
      setBooks(books.map((b) => (b.id === formData.id ? newBook : b)));
    } else {
      setBooks([newBook, ...books]);
    }

    setIsAddModalOpen(false);
    setSelectedBookDetail(null);
    resetForm();
  };

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
    });
    setIsbnInput('');
    stopCameraScan();
  };

  const handleEditBook = (book: BookItem) => {
    setFormData(book);
    setIsbnInput(book.isbn || '');
    setSelectedBookDetail(null);
    setIsAddModalOpen(true);
  };

  const handleDeleteBook = (id: string) => {
    if (confirm('Sei sicuro di voler eliminare questo libro?')) {
      setBooks(books.filter((b) => b.id !== id));
      if (selectedBookDetail?.id === id) setSelectedBookDetail(null);
    }
  };

  // Drag and Drop reordering per Sezione Letti
  const handleDragStart = (index: number) => {
    setDraggedIndex(index);
  };

  const handleDragOver = (e: React.DragEvent, index: number) => {
    e.preventDefault();
    if (draggedIndex === null || draggedIndex === index) return;

    const readOnlyBooks = books.filter((b) => b.isRead);
    const itemToMove = readOnlyBooks[draggedIndex];
    
    const updated = [...books];
    const sourceGlobalIdx = updated.findIndex((b) => b.id === itemToMove.id);
    const targetGlobalIdx = updated.findIndex((b) => b.id === readOnlyBooks[index].id);

    updated.splice(sourceGlobalIdx, 1);
    updated.splice(targetGlobalIdx, 0, itemToMove);

    setDraggedIndex(index);
    setBooks(updated);
  };

  const handleDragEnd = () => {
    setDraggedIndex(null);
  };

  // Esportazioni
  const exportToExcel = () => {
    const dataToExport = books.map((b) => ({
      Titolo: b.title,
      Autore: b.author,
      'Paese di Pubblicazione': b.publishCountry || '-',
      Classico: b.isClassic ? 'Sì' : 'No',
      Stato: b.isRead ? 'Letto' : 'In Biblioteca',
      Formato: b.format === 'cartaceo' ? 'Cartaceo' : b.format === 'ebook' ? 'eBook' : 'eBook + Cartaceo',
      Editore: b.publisher,
      'Anno Pubblicazione': b.publishYear,
      Genere: b.genre,
      'Serie / Tag': b.seriesTag,
      Volume: b.volume || '-',
      Pagine: b.pages,
      'Mese e Anno di Lettura': b.readMonthYear,
      ISBN: b.isbn,
    }));

    const worksheet = XLSX.utils.json_to_sheet(dataToExport);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Biblioteca');
    XLSX.writeFile(workbook, 'La_Mia_Biblioteca.xlsx');
  };

  const exportToPDF = () => {
    const doc = new jsPDF();
    doc.text('La Mia Biblioteca - Report', 14, 15);
    const tableData = books.map((b) => [
      b.title,
      b.author,
      b.publishCountry || '-',
      b.isClassic ? 'Sì' : 'No',
      b.isRead ? 'Letto' : 'In Libreria',
      b.format === 'cartaceo' ? 'Cartaceo' : b.format === 'ebook' ? 'eBook' : 'eBook + Cartaceo',
      b.genre || '-',
      b.readMonthYear || '-',
    ]);
    autoTable(doc, {
      head: [['Titolo', 'Autore', 'Paese', 'Classico', 'Stato', 'Formato', 'Genere', 'Data Lettura']],
      body: tableData,
      startY: 20,
    });
    doc.save('La_Mia_Biblioteca.pdf');
  };

  const exportBackup = () => {
    const dataStr = 'data:text/json;charset=utf-8,' + encodeURIComponent(JSON.stringify(books));
    const downloadAnchor = document.createElement('a');
    downloadAnchor.setAttribute('href', dataStr);
    downloadAnchor.setAttribute('download', 'backup_libreria.json');
    document.body.appendChild(downloadAnchor);
    downloadAnchor.click();
    downloadAnchor.remove();
  };

  const handleImportBackup = (e: React.ChangeEvent<HTMLInputElement>) => {
    const fileReader = new FileReader();
    if (e.target.files && e.target.files[0]) {
      fileReader.readAsText(e.target.files[0], 'UTF-8');
      fileReader.onload = (event) => {
        try {
          const parsed = JSON.parse(event.target?.result as string);
          if (Array.isArray(parsed)) {
            setBooks(parsed);
            alert('Backup ripristinato con successo!');
          }
        } catch (err) {
          alert('File di backup non valido.');
        }
      };
    }
  };

  const handleClearAll = () => {
    if (confirm('ATTENZIONE: Verranno cancellati TUTTI i libri salvati. Procedere?')) {
      setBooks([]);
      localStorage.removeItem('ios_library_books_v5');
    }
  };

  // Libri Letti Filtri
  const readBooksFiltered = books
    .filter((b) => b.isRead)
    .filter((b) => (filterGenre === 'all' ? true : b.genre === filterGenre))
    .filter((b) => {
      if (filterFormat === 'all') return true;
      if (filterFormat === 'cartaceo') return b.format === 'cartaceo';
      if (filterFormat === 'ebook') return b.format === 'ebook' || b.format === 'ebook_and_paper';
      return true;
    })
    .filter((b) => {
      if (filterYear === 'all') return true;
      return b.readYear?.toString() === filterYear || b.readMonthYear?.includes(filterYear);
    })
    .filter((b) =>
      searchQuery === ''
        ? true
        : b.title.toLowerCase().includes(searchQuery.toLowerCase()) ||
          b.author.toLowerCase().includes(searchQuery.toLowerCase())
    );

  // Raggruppamento Autori
  const authorsMap = books.reduce((acc, book) => {
    const authorName = book.author.trim() || 'Autore Sconosciuto';
    if (!acc[authorName]) acc[authorName] = [];
    acc[authorName].push(book);
    return acc;
  }, {} as Record<string, BookItem[]>);

  const sortedAuthors = Object.keys(authorsMap).sort((a, b) => a.localeCompare(b));

  // Generi Raggruppati
  const genresMap = books.reduce((acc, book) => {
    const genreName = book.genre?.trim() || 'Generico / Altro';
    if (!acc[genreName]) acc[genreName] = [];
    acc[genreName].push(book);
    return acc;
  }, {} as Record<string, BookItem[]>);

  const sortedGenres = Object.keys(genresMap).sort((a, b) => a.localeCompare(b));

  const availableGenres = Array.from(new Set(books.map((b) => b.genre).filter(Boolean)));
  const availableYears = Array.from(
    new Set(
      books
        .map((b) => {
          if (b.readYear) return b.readYear.toString();
          const match = b.readMonthYear?.match(/\d{4}/);
          return match ? match[0] : null;
        })
        .filter(Boolean)
    )
  );

  const formatLabel = (format: string) => {
    if (format === 'cartaceo') return 'Cartaceo';
    if (format === 'ebook') return 'eBook';
    if (format === 'ebook_and_paper') return 'eBook + Cartaceo';
    return format;
  };

  const renderAuthorGroup = (bookList: BookItem[]) => {
    const map = bookList.reduce((acc, book) => {
      const author = book.author.trim() || 'Autore Sconosciuto';
      if (!acc[author]) acc[author] = [];
      acc[author].push(book);
      return acc;
    }, {} as Record<string, BookItem[]>);

    return Object.keys(map).sort().map((author) => (
      <div key={author} className="space-y-2 pt-2">
        <h3 className="font-serif font-bold text-sm text-amber-950 border-b border-amber-900/10 pb-1">
          {author}
        </h3>
        <div className="grid grid-cols-2 gap-3">
          {map[author]
            .sort((a, b) => (a.volume || '').localeCompare(b.volume || ''))
            .map((book) => (
              <div
                key={book.id}
                onClick={() => setSelectedBookDetail(book)}
                className="bg-[#FFFDF9] p-3 rounded-2xl shadow-sm border border-amber-900/10 flex flex-col cursor-pointer"
              >
                <div className="w-full h-36 bg-amber-100/40 rounded-xl overflow-hidden relative mb-2 border border-amber-900/10">
                  {book.coverUrl ? (
                    <img src={book.coverUrl} alt={book.title} className="w-full h-full object-cover" />
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
                      Vol. {book.volume}
                    </div>
                  )}
                </div>
                <h4 className="font-serif font-bold text-xs text-amber-950 line-clamp-2">{book.title}</h4>
                <p className="text-[10px] text-amber-800/60 mt-0.5">{book.publishYear ? `Anno: ${book.publishYear}` : ''}</p>
              </div>
            ))}
        </div>
      </div>
    ));
  };

  return (
    <div className="min-h-screen bg-[#FBF9F5] text-amber-950 font-sans pb-24 select-none">
      {/* Header pulito: tasto + presente SOLO nella seconda tab ("read") */}
      <header className="sticky top-0 z-20 bg-[#FBF9F5]/90 backdrop-blur-md border-b border-amber-900/10 px-5 py-3.5 flex justify-between items-center">
        <div>
          <span className="text-[11px] font-bold text-amber-800/70 uppercase tracking-widest block">
            {activeTab === 'home' && 'La Mia Collezione'}
            {activeTab === 'read' && 'Cronologia Letture'}
            {activeTab === 'authors' && 'Catalogo Autori'}
            {activeTab === 'settings' && 'Gestione Dati'}
          </span>
          <h1 className="text-2xl font-serif font-extrabold tracking-tight text-amber-950">
            {activeTab === 'home' && 'Home'}
            {activeTab === 'read' && 'Libri Letti'}
            {activeTab === 'authors' && (selectedAuthor ? selectedAuthor : 'Autori')}
            {activeTab === 'settings' && 'Impostazioni'}
          </h1>
        </div>

        {activeTab === 'read' && (
          <button
            onClick={() => {
              resetForm();
              setIsAddModalOpen(true);
            }}
            className="w-10 h-10 bg-amber-800 text-amber-50 rounded-full flex items-center justify-center shadow-md active:scale-95 transition-transform"
          >
            <Plus className="w-5 h-5 stroke-[2.5]" />
          </button>
        )}
      </header>

      {/* VISTA 1: HOME */}
      {activeTab === 'home' && (
        <div className="p-4 space-y-5 max-w-lg mx-auto">
          {homeSubView !== 'none' ? (
            <div className="space-y-4">
              <button
                onClick={() => {
                  setHomeSubView('none');
                  setSelectedGenreHome(null);
                }}
                className="text-xs font-bold text-amber-800 flex items-center gap-1"
              >
                ← Torna alla Home
              </button>

              {homeSubView === 'classics' && (
                <div className="space-y-3">
                  <div className="flex items-center gap-2">
                    <Sparkles className="w-5 h-5 text-amber-600" />
                    <h2 className="text-lg font-serif font-bold text-amber-950">I Miei Classici</h2>
                  </div>
                  {renderAuthorGroup(books.filter((b) => b.isClassic))}
                </div>
              )}

              {homeSubView === 'genres' && (
                <div className="space-y-4">
                  {!selectedGenreHome ? (
                    <div className="space-y-2">
                      <h2 className="text-lg font-serif font-bold text-amber-950 mb-2">Generi nella Biblioteca</h2>
                      {sortedGenres.map((genre) => (
                        <div
                          key={genre}
                          onClick={() => setSelectedGenreHome(genre)}
                          className="bg-[#FFFDF9] p-4 rounded-2xl shadow-sm border border-amber-900/10 flex justify-between items-center cursor-pointer"
                        >
                          <div>
                            <h3 className="font-serif font-bold text-base text-amber-950">{genre}</h3>
                            <p className="text-xs text-amber-800/60 font-medium">
                              {genresMap[genre].length} {genresMap[genre].length === 1 ? 'libro' : 'libri'}
                            </p>
                          </div>
                          <ChevronRight className="w-5 h-5 text-amber-800/30" />
                        </div>
                      ))}
                    </div>
                  ) : (
                    <div className="space-y-3">
                      <button
                        onClick={() => setSelectedGenreHome(null)}
                        className="text-xs font-bold text-amber-800 mb-1 block"
                      >
                        ← Tutti i generi
                      </button>
                      <h2 className="text-lg font-serif font-bold text-amber-950">{selectedGenreHome}</h2>
                      {renderAuthorGroup(genresMap[selectedGenreHome])}
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
                      Anno {currentYearNum}
                    </span>
                    <span className="text-2xl font-serif font-bold text-amber-950">
                      {readThisYearCount} <span className="text-sm font-sans font-normal text-amber-800/70">libri letti</span>
                    </span>
                  </div>
                </div>

                <div className="pt-2.5 border-t border-amber-900/10 flex items-center justify-center text-xs font-semibold text-amber-900/80">
                  <span>{readThisYearCartacei} Cartacei e {readThisYearEbook} eBook</span>
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3.5">
                <div className="bg-[#FFFDF9] p-4 rounded-3xl shadow-sm border border-amber-900/10 flex flex-col justify-between min-h-[9rem]">
                  <div className="w-10 h-10 rounded-2xl bg-amber-100/70 text-amber-900 flex items-center justify-center">
                    <BookOpen className="w-5 h-5" />
                  </div>
                  <div className="mt-2">
                    <span className="text-3xl font-serif font-black text-amber-950 leading-none">{totalBooks}</span>
                    <span className="text-xs font-bold text-amber-800/60 uppercase tracking-wider block mt-1">
                      In Biblioteca
                    </span>
                    <div className="mt-2 pt-2 border-t border-amber-900/10 text-xs font-semibold text-amber-900/80">
                      {totalCartacei} Cartacei e {totalEbook} eBook
                    </div>
                  </div>
                </div>

                <div className="bg-[#FFFDF9] p-4 rounded-3xl shadow-sm border border-amber-900/10 flex flex-col justify-between min-h-[9rem]">
                  <div className="w-10 h-10 rounded-2xl bg-emerald-100/70 text-emerald-800 flex items-center justify-center">
                    <CheckCircle2 className="w-5 h-5" />
                  </div>
                  <div className="mt-2">
                    <span className="text-3xl font-serif font-black text-amber-950 leading-none">{readBooksCount}</span>
                    <span className="text-xs font-bold text-amber-800/60 uppercase tracking-wider block mt-1">
                      Libri Letti
                    </span>
                    <div className="mt-2 pt-2 border-t border-amber-900/10 text-xs font-semibold text-amber-900/80">
                      {readCartacei} Cartacei e {readEbook} eBook
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
                    onClick={() => setHomeSubView('classics')}
                    className="bg-gradient-to-br from-amber-700 to-amber-900 text-amber-50 p-4 rounded-3xl shadow-md cursor-pointer active:scale-95 transition-transform flex flex-col justify-between h-32 relative overflow-hidden"
                  >
                    <Sparkles className="w-6 h-6 text-amber-200" />
                    <div>
                      <span className="text-lg font-serif font-bold block">Classici</span>
                      <span className="text-xs text-amber-200/80 font-medium">
                        {books.filter((b) => b.isClassic).length} libri conservati
                      </span>
                    </div>
                  </div>

                  <div
                    onClick={() => setHomeSubView('genres')}
                    className="bg-gradient-to-br from-stone-800 to-amber-950 text-amber-50 p-4 rounded-3xl shadow-md cursor-pointer active:scale-95 transition-transform flex flex-col justify-between h-32 relative overflow-hidden"
                  >
                    <Layers className="w-6 h-6 text-amber-300" />
                    <div>
                      <span className="text-lg font-serif font-bold block">Generi</span>
                      <span className="text-xs text-amber-200/80 font-medium">
                        {sortedGenres.length} categorie
                      </span>
                    </div>
                  </div>
                </div>
              </div>
            </>
          )}
        </div>
      )}

      {/* VISTA 2: LIBRI LETTI */}
      {activeTab === 'read' && (
        <div className="p-4 space-y-4 max-w-lg mx-auto">
          <div className="space-y-2.5">
            <div className="relative">
              <Search className="w-4 h-4 absolute left-3.5 top-3 text-amber-800/40" />
              <input
                type="text"
                placeholder="Cerca nei libri letti..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="w-full pl-9 pr-4 py-2.5 bg-[#FFFDF9] border border-amber-900/10 rounded-2xl text-xs font-medium focus:outline-none focus:ring-2 focus:ring-amber-800/20"
              />
            </div>

            <div className="flex gap-2 overflow-x-auto pb-1 text-xs">
              <select
                value={filterGenre}
                onChange={(e) => setFilterGenre(e.target.value)}
                className="bg-[#FFFDF9] border border-amber-900/10 rounded-xl px-3 py-2 font-medium text-amber-950"
              >
                <option value="all">Tutti i generi</option>
                {availableGenres.map((g) => (
                  <option key={g} value={g}>
                    {g}
                  </option>
                ))}
              </select>

              <select
                value={filterFormat}
                onChange={(e) => setFilterFormat(e.target.value)}
                className="bg-[#FFFDF9] border border-amber-900/10 rounded-xl px-3 py-2 font-medium text-amber-950"
              >
                <option value="all">Tutti i formati</option>
                <option value="cartaceo">Cartaceo</option>
                <option value="ebook">eBook</option>
              </select>

              <select
                value={filterYear}
                onChange={(e) => setFilterYear(e.target.value)}
                className="bg-[#FFFDF9] border border-amber-900/10 rounded-xl px-3 py-2 font-medium text-amber-950"
              >
                <option value="all">Tutti gli anni</option>
                {availableYears.map((y) => (
                  <option key={y} value={y!}>
                    {y}
                  </option>
                ))}
              </select>
            </div>
          </div>

          <p className="text-[11px] text-amber-800/50 font-medium px-1">
            Trascina tramite l'icona a sinistra per riordinare l'elenco dei libri letti.
          </p>

          <div className="space-y-3">
            {readBooksFiltered.length === 0 ? (
              <div className="text-center py-12 bg-[#FFFDF9] rounded-3xl border border-dashed border-amber-900/20">
                <BookOpen className="w-8 h-8 text-amber-800/30 mx-auto mb-2" />
                <p className="text-xs text-amber-800/60 font-medium">Nessun libro letto trovato.</p>
              </div>
            ) : (
              readBooksFiltered.map((book, index) => (
                <div
                  key={book.id}
                  draggable
                  onDragStart={() => handleDragStart(index)}
                  onDragOver={(e) => handleDragOver(e, index)}
                  onDragEnd={handleDragEnd}
                  className={`bg-[#FFFDF9] p-3 rounded-2xl shadow-sm border border-amber-900/10 flex items-center gap-3 transition-colors ${
                    draggedIndex === index ? 'bg-amber-100/50' : 'hover:bg-amber-50/50'
                  }`}
                >
                  <div className="cursor-grab active:cursor-grabbing text-amber-800/30 hover:text-amber-800">
                    <GripVertical className="w-5 h-5" />
                  </div>

                  <div
                    onClick={() => setSelectedBookDetail(book)}
                    className="w-13 h-19 bg-amber-100/40 rounded-lg overflow-hidden flex-shrink-0 relative border border-amber-900/10 cursor-pointer"
                  >
                    {book.coverUrl ? (
                      <img src={book.coverUrl} alt={book.title} className="w-full h-full object-cover" />
                    ) : (
                      <div className="w-full h-full flex items-center justify-center text-amber-800/30">
                        <Book className="w-5 h-5" />
                      </div>
                    )}
                  </div>

                  <div
                    onClick={() => setSelectedBookDetail(book)}
                    className="flex-1 min-w-0 cursor-pointer"
                  >
                    <h3 className="font-serif font-bold text-sm text-amber-950 truncate">{book.title}</h3>
                    <p className="text-xs text-amber-800/70 truncate">{book.author}</p>
                    {book.rating && (
                      <div className="flex text-amber-500 mt-1">
                        {[...Array(5)].map((_, i) => (
                          <Star
                            key={i}
                            className={`w-3 h-3 ${
                              i < book.rating! ? 'fill-amber-500' : 'text-amber-200'
                            }`}
                          />
                        ))}
                      </div>
                    )}
                    <p className="text-[10px] font-semibold text-amber-800/50 mt-1">
                      Letto: <span className="text-amber-950">{book.readMonthYear || 'Data non specificata'}</span>
                    </p>
                  </div>
                  <ChevronRight className="w-4 h-4 text-amber-800/30" />
                </div>
              ))
            )}
          </div>
        </div>
      )}

      {/* VISTA 3: AUTORI */}
      {activeTab === 'authors' && (
        <div className="p-4 space-y-4 max-w-lg mx-auto">
          {!selectedAuthor ? (
            <div className="space-y-2">
              {sortedAuthors.length === 0 ? (
                <div className="text-center py-12 bg-[#FFFDF9] rounded-3xl border border-dashed border-amber-900/20">
                  <Users className="w-8 h-8 text-amber-800/30 mx-auto mb-2" />
                  <p className="text-xs text-amber-800/60 font-medium">Nessun autore presente.</p>
                </div>
              ) : (
                sortedAuthors.map((author) => {
                  const authorBooks = authorsMap[author];
                  return (
                    <div
                      key={author}
                      onClick={() => setSelectedAuthor(author)}
                      className="bg-[#FFFDF9] p-4 rounded-2xl shadow-sm border border-amber-900/10 flex justify-between items-center cursor-pointer"
                    >
                      <div>
                        <h3 className="font-serif font-bold text-base text-amber-950">{author}</h3>
                        <p className="text-xs text-amber-800/60 font-medium">
                          {authorBooks.length} {authorBooks.length === 1 ? 'libro' : 'libri'}
                        </p>
                      </div>
                      <ChevronRight className="w-5 h-5 text-amber-800/30" />
                    </div>
                  );
                })
              )}
            </div>
          ) : (
            <div className="space-y-4">
              <button
                onClick={() => setSelectedAuthor(null)}
                className="text-xs font-bold text-amber-800 flex items-center gap-1 mb-2"
              >
                ← Torna alla lista autori
              </button>

              <div className="grid grid-cols-2 gap-3.5">
                {authorsMap[selectedAuthor]?.map((book) => (
                  <div
                    key={book.id}
                    onClick={() => setSelectedBookDetail(book)}
                    className="bg-[#FFFDF9] p-3 rounded-2xl shadow-sm border border-amber-900/10 flex flex-col cursor-pointer relative"
                  >
                    <div className="w-full h-44 bg-amber-100/40 rounded-xl overflow-hidden relative border border-amber-900/10 mb-2">
                      {book.coverUrl ? (
                        <img src={book.coverUrl} alt={book.title} className="w-full h-full object-cover" />
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
                        {formatLabel(book.format)}
                      </div>

                      {book.volume && (
                        <div className="absolute top-2 left-2 bg-amber-800 text-amber-50 text-[9px] font-bold px-2 py-0.5 rounded-full shadow">
                          Vol. {book.volume}
                        </div>
                      )}
                    </div>

                    <h4 className="font-serif font-bold text-xs text-amber-950 line-clamp-2">{book.title}</h4>
                    <p className="text-[10px] font-medium text-amber-800/60 mt-1">
                      {book.publishYear ? `Anno: ${book.publishYear}` : ''}
                    </p>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      )}

      {/* VISTA 4: SETTINGS */}
      {activeTab === 'settings' && (
        <div className="p-4 space-y-5 max-w-lg mx-auto">
          <div className="bg-[#FFFDF9] rounded-3xl p-4 shadow-sm border border-amber-900/10 space-y-3">
            <h2 className="text-xs font-bold text-amber-800/60 uppercase tracking-wider">
              Esporta la tua Biblioteca
            </h2>

            <button
              onClick={exportToExcel}
              className="w-full p-3 bg-emerald-50 text-emerald-900 rounded-2xl font-bold text-xs flex items-center gap-3 border border-emerald-200"
            >
              <FileSpreadsheet className="w-5 h-5 text-emerald-700" /> Esporta in Foglio Excel (.xlsx)
            </button>

            <button
              onClick={exportToPDF}
              className="w-full p-3 bg-rose-50 text-rose-900 rounded-2xl font-bold text-xs flex items-center gap-3 border border-rose-200"
            >
              <FileText className="w-5 h-5 text-rose-700" /> Esporta Report PDF
            </button>
          </div>

          <div className="bg-[#FFFDF9] rounded-3xl p-4 shadow-sm border border-amber-900/10 space-y-3">
            <h2 className="text-xs font-bold text-amber-800/60 uppercase tracking-wider">
              Backup & Ripristino Dati
            </h2>

            <button
              onClick={exportBackup}
              className="w-full p-3 bg-amber-100/60 text-amber-950 rounded-2xl font-bold text-xs flex items-center gap-3 border border-amber-200"
            >
              <Download className="w-5 h-5 text-amber-800" /> Salva Backup Dati (JSON)
            </button>

            <label className="w-full p-3 bg-stone-100/80 text-stone-900 rounded-2xl font-bold text-xs flex items-center gap-3 border border-stone-200 cursor-pointer">
              <Upload className="w-5 h-5 text-stone-700" /> Ripristina Backup da File
              <input type="file" accept=".json" onChange={handleImportBackup} className="hidden" />
            </label>
          </div>

          <div className="bg-[#FFFDF9] rounded-3xl p-4 shadow-sm border border-amber-900/10 space-y-3">
            <h2 className="text-xs font-bold text-rose-700 uppercase tracking-wider">Zona Pericolo</h2>
            <button
              onClick={handleClearAll}
              className="w-full p-3 bg-rose-700 text-amber-50 rounded-2xl font-bold text-xs flex items-center justify-center gap-2"
            >
              <Trash2 className="w-4 h-4" /> Cancella Intera Biblioteca
            </button>
          </div>
        </div>
      )}

      {/* MODALE SCHEDA LIBRO DETTAGLIO */}
      {selectedBookDetail && (
        <div className="fixed inset-0 z-50 bg-amber-950/40 backdrop-blur-sm flex items-end sm:items-center justify-center p-0 sm:p-4">
          <div className="bg-[#FFFDF9] w-full max-w-lg rounded-t-3xl sm:rounded-3xl max-h-[90vh] overflow-y-auto p-6 space-y-4 shadow-2xl relative border border-amber-900/10">
            <button
              onClick={() => setSelectedBookDetail(null)}
              className="absolute top-4 right-4 p-2 bg-amber-100/50 rounded-full text-amber-900"
            >
              <X className="w-5 h-5" />
            </button>

            <div className="flex gap-4 items-start pt-2">
              <div className="w-24 h-36 bg-amber-100/40 rounded-xl overflow-hidden border border-amber-900/10 flex-shrink-0">
                {selectedBookDetail.coverUrl ? (
                  <img
                    src={selectedBookDetail.coverUrl}
                    alt={selectedBookDetail.title}
                    className="w-full h-full object-cover"
                  />
                ) : (
                  <div className="w-full h-full flex items-center justify-center text-amber-800/30">
                    <Book className="w-8 h-8" />
                  </div>
                )}
              </div>

              <div className="flex-1">
                <h2 className="text-lg font-serif font-bold text-amber-950 leading-tight">{selectedBookDetail.title}</h2>
                <p className="text-xs font-medium text-amber-800/70 mt-1">{selectedBookDetail.author}</p>
                <div className="flex flex-wrap gap-2 mt-3">
                  <span className="text-[10px] px-2.5 py-1 rounded-full bg-amber-100/70 text-amber-900 font-bold">
                    {formatLabel(selectedBookDetail.format)}
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
                <span className="block text-amber-800/50 font-semibold text-[10px]">Paese di Pubblicazione:</span>
                <span className="font-bold">{selectedBookDetail.publishCountry || '-'}</span>
              </div>
              <div>
                <span className="block text-amber-800/50 font-semibold text-[10px]">Editore:</span>
                <span className="font-bold">{selectedBookDetail.publisher || '-'}</span>
              </div>
              <div>
                <span className="block text-amber-800/50 font-semibold text-[10px]">Anno Pubblicazione:</span>
                <span className="font-bold">{selectedBookDetail.publishYear || '-'}</span>
              </div>
              <div>
                <span className="block text-amber-800/50 font-semibold text-[10px]">Pagine:</span>
                <span className="font-bold">{selectedBookDetail.pages || '-'}</span>
              </div>
              <div>
                <span className="block text-amber-800/50 font-semibold text-[10px]">Genere:</span>
                <span className="font-bold">{selectedBookDetail.genre || '-'}</span>
              </div>
              <div>
                <span className="block text-amber-800/50 font-semibold text-[10px]">Serie / Tag:</span>
                <span className="font-bold">{selectedBookDetail.seriesTag || '-'}</span>
              </div>
              <div>
                <span className="block text-amber-800/50 font-semibold text-[10px]">Volume:</span>
                <span className="font-bold">{selectedBookDetail.volume || '-'}</span>
              </div>
              {selectedBookDetail.isRead && (
                <div className="col-span-2 border-t border-amber-900/10 pt-2">
                  <span className="block text-amber-800/50 font-semibold text-[10px]">Mese & Anno di Lettura:</span>
                  <span className="font-bold">{selectedBookDetail.readMonthYear || '-'}</span>
                </div>
              )}
            </div>

            <div className="grid grid-cols-2 gap-3 pt-2">
              <button
                onClick={() => handleEditBook(selectedBookDetail)}
                className="py-3 bg-amber-800 text-amber-50 rounded-xl font-bold text-xs flex items-center justify-center gap-2"
              >
                <Edit className="w-4 h-4" /> Modifica
              </button>

              <button
                onClick={() => handleDeleteBook(selectedBookDetail.id)}
                className="py-3 bg-rose-100 text-rose-800 rounded-xl font-bold text-xs flex items-center justify-center gap-2 border border-rose-200"
              >
                <Trash2 className="w-4 h-4" /> Elimina
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODALE AGGIUNTA / MODIFICA LIBRO */}
      {isAddModalOpen && (
        <div className="fixed inset-0 z-50 bg-amber-950/40 backdrop-blur-sm flex items-end sm:items-center justify-center p-0 sm:p-4">
          <div className="bg-[#FFFDF9] w-full max-w-lg rounded-t-3xl sm:rounded-3xl max-h-[90vh] overflow-y-auto p-6 space-y-4 shadow-2xl relative border border-amber-900/10">
            <div className="flex justify-between items-center border-b border-amber-900/10 pb-3">
              <h2 className="text-base font-serif font-bold text-amber-950">
                {formData.id ? 'Modifica Libro' : 'Aggiungi Nuovo Libro'}
              </h2>
              <button
                onClick={() => {
                  stopCameraScan();
                  setIsAddModalOpen(false);
                }}
                className="p-2 text-amber-800/40"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="bg-amber-100/50 p-3.5 rounded-2xl space-y-3 border border-amber-900/10">
              <div className="flex justify-between items-center">
                <span className="text-[10px] font-bold text-amber-900 uppercase tracking-wider">
                  Compilazione Automatica (Codice ISBN)
                </span>
                <Scan className="w-4 h-4 text-amber-800" />
              </div>

              {isScanningCamera ? (
                <div className="relative rounded-xl overflow-hidden bg-black aspect-video flex items-center justify-center">
                  <video ref={videoRef} className="w-full h-full object-cover" />
                  <button
                    type="button"
                    onClick={stopCameraScan}
                    className="absolute bottom-2 px-3 py-1 bg-rose-600 text-white text-[11px] font-bold rounded-lg shadow"
                  >
                    Chiudi Fotocamera
                  </button>
                </div>
              ) : (
                <button
                  type="button"
                  onClick={startCameraScan}
                  className="w-full py-2.5 bg-amber-800 text-amber-50 rounded-xl font-bold text-xs flex items-center justify-center gap-2 shadow-sm active:scale-95 transition-transform"
                >
                  <Camera className="w-4 h-4" /> Scansiona con Fotocamera
                </button>
              )}

              <div className="flex gap-2">
                <input
                  type="text"
                  placeholder="O inserisci codice ISBN manualmente"
                  value={isbnInput}
                  onChange={(e) => setIsbnInput(e.target.value)}
                  className="flex-1 p-2.5 bg-[#FFFDF9] border border-amber-900/10 rounded-xl text-xs font-medium focus:outline-none"
                />
                <button
                  type="button"
                  onClick={() => handleSearchBookByISBN()}
                  disabled={isSearchingIsbn}
                  className="px-4 bg-amber-800 text-amber-50 rounded-xl text-xs font-bold active:scale-95 transition-transform"
                >
                  {isSearchingIsbn ? '...' : 'Cerca'}
                </button>
              </div>
            </div>

            <form onSubmit={handleSaveBook} className="space-y-3 text-xs">
              <input
                type="text"
                placeholder="Titolo *"
                value={formData.title}
                onChange={(e) => setFormData({ ...formData, title: e.target.value })}
                className="w-full p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl font-medium focus:outline-none"
                required
              />

              <div className="grid grid-cols-2 gap-2">
                <input
                  type="text"
                  placeholder="Autore *"
                  value={formData.author}
                  onChange={(e) => setFormData({ ...formData, author: e.target.value })}
                  className="p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl font-medium focus:outline-none"
                  required
                />
                <input
                  type="text"
                  placeholder="Paese di pubblicazione"
                  value={formData.publishCountry || ''}
                  onChange={(e) => setFormData({ ...formData, publishCountry: e.target.value })}
                  className="p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl font-medium focus:outline-none"
                />
              </div>

              <input
                type="text"
                placeholder="URL Copertina Immagine"
                value={formData.coverUrl}
                onChange={(e) => setFormData({ ...formData, coverUrl: e.target.value })}
                className="w-full p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl font-medium focus:outline-none"
              />

              <div className="grid grid-cols-2 gap-2">
                <input
                  type="text"
                  placeholder="Editore"
                  value={formData.publisher}
                  onChange={(e) => setFormData({ ...formData, publisher: e.target.value })}
                  className="p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl font-medium focus:outline-none"
                />
                <input
                  type="text"
                  placeholder="Anno Pubblicazione"
                  value={formData.publishYear}
                  onChange={(e) => setFormData({ ...formData, publishYear: e.target.value })}
                  className="p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl font-medium focus:outline-none"
                />
              </div>

              <div className="grid grid-cols-2 gap-2">
                <input
                  type="text"
                  placeholder="Genere"
                  value={formData.genre}
                  onChange={(e) => setFormData({ ...formData, genre: e.target.value })}
                  className="p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl font-medium focus:outline-none"
                />
                <input
                  type="text"
                  inputMode="numeric"
                  pattern="[0-9]*"
                  placeholder="Numero Pagine"
                  value={formData.pages || ''}
                  onChange={(e) => {
                    const val = e.target.value.replace(/\D/g, '');
                    setFormData({ ...formData, pages: val ? parseInt(val) : undefined });
                  }}
                  className="p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl font-medium focus:outline-none"
                />
              </div>

              <div className="grid grid-cols-2 gap-2">
                <input
                  type="text"
                  placeholder="Serie / Tag"
                  value={formData.seriesTag}
                  onChange={(e) => setFormData({ ...formData, seriesTag: e.target.value })}
                  className="p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl font-medium focus:outline-none"
                />
                <input
                  type="text"
                  placeholder="Volume"
                  value={formData.volume || ''}
                  onChange={(e) => setFormData({ ...formData, volume: e.target.value })}
                  className="p-3 bg-[#FFFDF9] border border-amber-900/10 rounded-xl font-medium focus:outline-none"
                />
              </div>

              <div className="flex items-center justify-between p-3 bg-amber-50/60 rounded-xl border border-amber-900/10">
                <span className="font-bold text-amber-950 flex items-center gap-1.5">
                  <Sparkles className="w-4 h-4 text-amber-600" /> È un Classico?
                </span>
                <input
                  type="checkbox"
                  checked={formData.isClassic}
                  onChange={(e) => setFormData({ ...formData, isClassic: e.target.checked })}
                  className="w-5 h-5 accent-amber-800"
                />
              </div>

              <div className="p-3 bg-amber-50/60 rounded-xl space-y-2 border border-amber-900/10">
                <span className="font-bold text-amber-950 block">Formato:</span>
                <div className="flex gap-2">
                  <button
                    type="button"
                    onClick={() => setFormData({ ...formData, format: 'cartaceo' })}
                    className={`flex-1 py-2 rounded-lg text-xs font-bold transition-colors ${
                      formData.format === 'cartaceo'
                        ? 'bg-amber-800 text-amber-50'
                        : 'bg-[#FFFDF9] border border-amber-900/10 text-amber-900'
                    }`}
                  >
                    Cartaceo
                  </button>
                  <button
                    type="button"
                    onClick={() => setFormData({ ...formData, format: 'ebook' })}
                    className={`flex-1 py-2 rounded-lg text-xs font-bold transition-colors ${
                      formData.format === 'ebook' || formData.format === 'ebook_and_paper'
                        ? 'bg-amber-800 text-amber-50'
                        : 'bg-[#FFFDF9] border border-amber-900/10 text-amber-900'
                    }`}
                  >
                    eBook
                  </button>
                </div>

                {(formData.format === 'ebook' || formData.format === 'ebook_and_paper') && (
                  <label className="flex items-center gap-2 pt-1 cursor-pointer">
                    <input
                      type="checkbox"
                      checked={formData.format === 'ebook_and_paper'}
                      onChange={(e) =>
                        setFormData({
                          ...formData,
                          format: e.target.checked ? 'ebook_and_paper' : 'ebook',
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

              <div className="flex items-center justify-between p-3 bg-amber-50/60 rounded-xl border border-amber-900/10">
                <span className="font-bold text-amber-950">Letto</span>
                <input
                  type="checkbox"
                  checked={formData.isRead}
                  onChange={(e) => setFormData({ ...formData, isRead: e.target.checked })}
                  className="w-5 h-5 accent-amber-800"
                />
              </div>

              {formData.isRead && (
                <div className="p-3 bg-amber-100/40 border border-amber-900/10 rounded-xl space-y-2">
                  <span className="font-bold text-amber-950 block text-[11px] uppercase tracking-wider">
                    Data di Lettura
                  </span>
                  <div className="grid grid-cols-2 gap-2">
                    <div>
                      <label className="block text-[10px] text-amber-800/70 font-semibold mb-1">Mese:</label>
                      <select
                        value={formData.readMonth || 'Gennaio'}
                        onChange={(e) => setFormData({ ...formData, readMonth: e.target.value })}
                        className="w-full p-2.5 bg-[#FFFDF9] border border-amber-900/10 rounded-xl font-semibold text-amber-950 focus:outline-none"
                      >
                        {MONTHS.map((m) => (
                          <option key={m} value={m}>
                            {m}
                          </option>
                        ))}
                      </select>
                    </div>

                    <div>
                      <label className="block text-[10px] text-amber-800/70 font-semibold mb-1">Anno:</label>
                      <select
                        value={formData.readYear || currentYearNum}
                        onChange={(e) => setFormData({ ...formData, readYear: parseInt(e.target.value) })}
                        className="w-full p-2.5 bg-[#FFFDF9] border border-amber-900/10 rounded-xl font-semibold text-amber-950 focus:outline-none"
                      >
                        {yearsList.map((y) => (
                          <option key={y} value={y}>
                            {y}
                          </option>
                        ))}
                      </select>
                    </div>
                  </div>
                </div>
              )}

              <button
                type="submit"
                className="w-full py-3 bg-amber-800 text-amber-50 rounded-xl font-bold text-xs shadow-md active:scale-95 transition-transform"
              >
                {formData.id ? 'Aggiorna Libro' : 'Salva Libro'}
              </button>
            </form>
          </div>
        </div>
      )}

      {/* TAB BAR INFERIORE */}
      <nav className="fixed bottom-0 left-0 right-0 z-30 bg-[#FBF9F5]/90 backdrop-blur-md border-t border-amber-900/10 flex justify-around py-2.5 max-w-lg mx-auto">
        <button
          onClick={() => {
            setActiveTab('home');
            setSelectedAuthor(null);
            setHomeSubView('none');
          }}
          className={`flex flex-col items-center gap-1 ${
            activeTab === 'home' ? 'text-amber-800' : 'text-amber-900/40'
          }`}
        >
          <Home className="w-5 h-5" />
          <span className="text-[10px] font-bold">Home</span>
        </button>

        <button
          onClick={() => {
            setActiveTab('read');
            setSelectedAuthor(null);
          }}
          className={`flex flex-col items-center gap-1 ${
            activeTab === 'read' ? 'text-amber-800' : 'text-amber-900/40'
          }`}
        >
          <BookOpen className="w-5 h-5" />
          <span className="text-[10px] font-bold">Letti</span>
        </button>

        <button
          onClick={() => {
            setActiveTab('authors');
            setSelectedAuthor(null);
          }}
          className={`flex flex-col items-center gap-1 ${
            activeTab === 'authors' ? 'text-amber-800' : 'text-amber-900/40'
          }`}
        >
          <Users className="w-5 h-5" />
          <span className="text-[10px] font-bold">Autori</span>
        </button>

        <button
          onClick={() => {
            setActiveTab('settings');
            setSelectedAuthor(null);
          }}
          className={`flex flex-col items-center gap-1 ${
            activeTab === 'settings' ? 'text-amber-800' : 'text-amber-900/40'
          }`}
        >
          <Settings className="w-5 h-5" />
          <span className="text-[10px] font-bold">Settings</span>
        </button>
      </nav>
    </div>
  );
}