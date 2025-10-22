# SLIDE RELEVANCIA ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    # Adicionar um novo slide para 'Relevância'
    slide_layout = prs.slide_layouts[5]  # Escolhendo um layout de slide em branco
    slide = prs.slides.add_slide(slide_layout)

    # Caminho para a imagem de fundo do novo slide
    background_image_path = 'resources/template_pagina.png'

    # Adicionar a imagem de fundo do novo slide
    slide.shapes.add_picture(background_image_path, 0, 0, width=prs.slide_width, height=prs.slide_height)

    # Adicionar a primeira caixa de texto do novo slide
    textbox_first = slide.shapes.add_textbox(Cm(1.68), Cm(0.71), Cm(18), Cm(1.8))
    text_frame_first = textbox_first.text_frame
    p_first = text_frame_first.paragraphs[0]
    run_first = p_first.add_run()
    run_first.text = "Relevância"
    font_first = run_first.font
    font_first.size = Pt(30)
    font_first.bold = True
    font_first.name = 'Poppins'
    font_first.color.rgb = RGBColor(25, 46, 104)  # Cor do texto em azul escuro

    # Adicionar a segunda caixa de texto do novo slide
    textbox_second = slide.shapes.add_textbox(Cm(1.9), Cm(2.32), Cm(9.38), Cm(1.11))
    text_frame_second = textbox_second.text_frame
    p_second = text_frame_second.paragraphs[0]
    run_second = p_second.add_run()
    run_second.text = artist
    font_second = run_second.font
    font_second.size = Pt(20)
    font_second.bold = True
    font_second.name = 'Poppins'
    font_second.color.rgb = RGBColor(255, 255, 255)  # Cor do texto em branco
    fill = textbox_second.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(113, 55, 248)  # Cor de preenchimento em roxo

    # Adicionar a terceira caixa de texto do novo slide
    textbox_third = slide.shapes.add_textbox(Cm(1.9), Cm(3.94), Cm(9.38), Cm(9.49))
    text_frame_third = textbox_third.text_frame
    text_frame_third.word_wrap = True
    text_frame_third.auto_size = True
    p_third = text_frame_third.paragraphs[0]
    run_third = p_third.add_run()
    run_third.text = arquivo[arquivo['Tópicos'] == 'Relevância']['Textos'].iloc[0]
    font_third = run_third.font
    font_third.size = Pt(16)
    font_third.name = 'Poppins'
    font_third.color.rgb = RGBColor(0, 0, 0)  # Cor do texto em preto

    # Adicionar a imagem
    image_path = f'dados_full/{artista}/plots/Relevancia.jpg'
    slide.shapes.add_picture(image_path, Cm(11.74), Cm(4.42), width=Cm(21.69), height=Cm(11.84))

 #SLIDE RECEITA ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    # Adicionar um novo slide para 'Receita'
    slide_layout = prs.slide_layouts[5]  # Escolhendo um layout de slide em branco
    slide = prs.slides.add_slide(slide_layout)

    # Caminho para a imagem de fundo do novo slide
    background_image_path = 'resources/template_pagina.png'

    # Adicionar a imagem de fundo do novo slide
    slide.shapes.add_picture(background_image_path, 0, 0, width=prs.slide_width, height=prs.slide_height)

    # Adicionar a primeira caixa de texto do novo slide
    textbox_first = slide.shapes.add_textbox(Cm(1.68), Cm(0.71), Cm(18), Cm(1.8))
    text_frame_first = textbox_first.text_frame
    p_first = text_frame_first.paragraphs[0]
    run_first = p_first.add_run()
    run_first.text = "Receita"
    font_first = run_first.font
    font_first.size = Pt(30)
    font_first.bold = True
    font_first.name = 'Poppins'
    font_first.color.rgb = RGBColor(25, 46, 104)  # Cor do texto em azul escuro

    # Adicionar a segunda caixa de texto do novo slide
    textbox_second = slide.shapes.add_textbox(Cm(1.9), Cm(2.32), Cm(9.38), Cm(1.11))
    text_frame_second = textbox_second.text_frame
    p_second = text_frame_second.paragraphs[0]
    run_second = p_second.add_run()
    run_second.text = artist
    font_second = run_second.font
    font_second.size = Pt(20)
    font_second.bold = True
    font_second.name = 'Poppins'
    font_second.color.rgb = RGBColor(255, 255, 255)  # Cor do texto em branco
    fill = textbox_second.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(113, 55, 248)  # Cor de preenchimento em roxo

    # Adicionar a terceira caixa de texto do novo slide
    textbox_third = slide.shapes.add_textbox(Cm(1.9), Cm(3.94), Cm(9.38), Cm(9.49))
    text_frame_third = textbox_third.text_frame
    text_frame_third.word_wrap = True
    text_frame_third.auto_size = True
    p_third = text_frame_third.paragraphs[0]
    run_third = p_third.add_run()
    run_third.text = arquivo[arquivo['Tópicos'] == 'Receita']['Textos'].iloc[0]
    font_third = run_third.font
    font_third.size = Pt(16)
    font_third.name = 'Poppins'
    font_third.color.rgb = RGBColor(0, 0, 0)  # Cor do texto em preto

    # Adicionar a imagem
    image_path = f'dados_full/{artista}/plots/Receita.jpg'
    slide.shapes.add_picture(image_path, Cm(11.74), Cm(4.42), width=Cm(21.69), height=Cm(11.84))

#SLIDE CONCLUSAO ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    slide_layout = prs.slide_layouts[5]  # Escolhendo um layout de slide em branco
    slide = prs.slides.add_slide(slide_layout)

    # Caminho para a imagem de fundo do novo slide
    background_image_path_1 = 'resources/template_pagina.png'

    # Adicionar a imagem de fundo do novo slide
    slide.shapes.add_picture(background_image_path_1, 0, 0, width=prs.slide_width, height=prs.slide_height)

    # Adicionar a primeira caixa de texto do novo slide
    textbox_first = slide.shapes.add_textbox(Cm(1.68), Cm(0.71), Cm(18), Cm(1.8))
    text_frame_first = textbox_first.text_frame
    p_first = text_frame_first.paragraphs[0]
    run_first = p_first.add_run()
    run_first.text = "Conclusão"
    font_first = run_first.font
    font_first.size = Pt(30)
    font_first.bold = True
    font_first.name = 'Poppins'
    font_first.color.rgb = RGBColor(25, 46, 104)  # Cor do texto em formato RGB hexadecimal

    # Adicionar a segunda caixa de texto do novo slide
    textbox_second = slide.shapes.add_textbox(Cm(1.9), Cm(2.32), Cm(9.38), Cm(1.11))
    text_frame_second = textbox_second.text_frame
    p_second = text_frame_second.paragraphs[0]
    run_second = p_second.add_run()
    run_second.text = artist
    font_second = run_second.font
    font_second.size = Pt(20)
    font_second.bold = True
    font_second.name = 'Poppins'
    font_second.color.rgb = RGBColor(255, 255, 255)  # Cor do texto em branco
    # Definir cor de preenchimento da caixa de texto
    fill = textbox_second.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(113, 55, 248)  # Cor de preenchimento em formato RGB hexadecimal

    # Adicionar a terceira caixa de texto do novo slide
    textbox_third = slide.shapes.add_textbox(Cm(1.9), Cm(3.94), Cm(30.5), Cm(13.25))
    text_frame_third = textbox_third.text_frame
    text_frame_third.word_wrap = True
    text_frame_third.auto_size = True
    p_third = text_frame_third.paragraphs[0]
    run_third = p_third.add_run()
    run_third.text = arquivo[arquivo['Tópicos'] == 'Conclusão']['Textos'].iloc[0]
    font_third = run_third.font
    font_third.size = Pt(16)
    font_third.name = 'Poppins'
    font_third.color.rgb = RGBColor(0, 0, 0)  # Cor do texto em preto