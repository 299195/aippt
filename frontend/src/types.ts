export type StyleType = "management" | "technical";
export type RewriteAction = "concise" | "management" | "technical";
export type TemplateId = string;

export interface TemplateItem {
  id: TemplateId;
  name: string;
  subtitle: string;
  summary: string;
  preview_bg: string;
  preview_fg: string;
  preview_accent: string;
  preview_image_url?: string | null;
}

export interface Slide {
  page: number;
  title: string;
  bullets: string[];
  notes: string;
  slide_type?: string;
  evidence?: string[];
}

export interface JobDetail {
  job_id: string;
  status: string;
  style: StyleType;
  template_id: TemplateId;
  title: string;
  outline: string[];
  slides: Slide[];
  pptx_url: string | null;
  created_at: string;
}

export interface HistoryItem {
  job_id: string;
  title: string;
  style: StyleType;
  template_id: TemplateId;
  status: string;
  created_at: string;
}

export interface ModelConfig {
  provider: string;
  model: string;
  use_mock: boolean;
  configured: boolean;
  base_url: string;
}

