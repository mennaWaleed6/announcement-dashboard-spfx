export type Lang = "EN" | "AR";

export const UI = {
  EN: {
    HeaderTitle: "Title",
    HeaderDescription: "Description",
    HeaderCategory: "Category",
    HeaderPriority: "Priority",
    HeaderDueDate: "Due Date",
    HeaderAssignedTo: "Assigned To",
  },
  AR: {
    HeaderTitle: "العنوان",
    HeaderDescription: "الوصف",
    HeaderCategory: "الفئة",
    HeaderPriority: "الأولوية",
    HeaderDueDate: "تاريخ الاستحقاق",
    HeaderAssignedTo: "مُسنَد إلى",
  },
} as const;
