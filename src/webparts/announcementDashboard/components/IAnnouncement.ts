export interface IAnnouncement {
  Title_ar: string;
  Title_en: string;
  Description_ar: string;
  Description_en: string;
  category?: {
    Title: string;
    name_en?: string;
    name_ar?: string;
  };
  Priority: number;
  DueDate: string;
  AssignedTo?: {
    Title: string;
    EMail?: string;
  };
}
