from django.urls import path
from . import views

urlpatterns = [
    path('', views.home, name='home'),
    path('all-tools/', views.all_tools, name='all_tools'),
    
    # Conversion Tools
    path('pdf-to-excel/', views.pdf_to_excel, name='pdf_to_excel'),
    path('pdf-to-word/', views.pdf_to_word, name='pdf_to_word'),
    path('pdf-to-ppt/', views.pdf_to_ppt, name='pdf_to_ppt'),
    path('pdf-to-images/', views.pdf_to_images, name='pdf_to_images'),
    
    # Editing Tools
    path('merge-pdf/', views.merge_pdf, name='merge_pdf'),
    path('split-pdf/', views.split_pdf, name='split_pdf'),
    path('protect-pdf/', views.protect_pdf, name='protect_pdf'),
    path('unlock-pdf/', views.unlock_pdf, name='unlock_pdf'),
    path('image-to-pdf/', views.image_to_pdf, name='image_to_pdf'),
    path('edit-pdf/', views.edit_pdf, name='edit_pdf'),
    path('compress-pdf/', views.compress_pdf, name='compress_pdf'),
    path('rotate-pdf/', views.rotate_pdf, name='rotate_pdf'),
    path('add-watermark/', views.add_watermark, name='add_watermark'),
    path('remove-watermark/', views.remove_watermark, name='remove_watermark'),
    path('pdf-to-text/', views.pdf_to_text, name='pdf_to_text'),
    path('html-to-pdf/', views.html_to_pdf, name='html_to_pdf'),
    path('edit-metadata/', views.edit_metadata, name='edit_metadata'),
    path('rearrange-pdf/', views.rearrange_pdf, name='rearrange_pdf'),
    path('fill-pdf-form/', views.fill_pdf_form, name='fill_pdf_form'),
    path('batch-process/', views.batch_process, name='batch_process'),
    path('pdf-info/', views.pdf_info, name='pdf_info'),
    path('ocr-pdf/', views.ocr_pdf, name='ocr_pdf'),
    path('about/', views.about, name='about'),
    path('contact/', views.contact, name='contact'),
    path('careers/', views.careers, name='careers'),
    path('blog/', views.blog, name='blog'),
    path('affiliate/', views.affiliate, name='affiliate'),
    path('privacy-policy/', views.privacy_policy, name='privacy_policy'),
    path('terms-of-service/', views.terms_of_service, name='terms_of_service'),
    path('disclaimer/', views.disclaimer, name='disclaimer'),
    path('cookie-policy/', views.cookie_policy, name='cookie_policy'),
    path('gdpr/', views.gdpr, name='gdpr'),
]