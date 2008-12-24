create or replace function DiasLetivos 
  (data_inicio  in varchar2,
   data_fim     in varchar2,
   cliente      in number,
   calendario   in number default null
  ) return number is
/**********************************************************************************
* Nome      : DiasLetivos
* Finalidade: Calcular o nº de dias letivos entre duas datas
* Autor     : Alexandre Vinhadelli Papadópolis
* Data      :  16/09/2008, 13:00
* Parâmetros:
*    data_inicio   : data inicial
*    data_fim      : data final
*    cliente       : indica o cliente
*    calendario    : indica o calendário do cliente
* Retorno: resultado calculado para o número de dias entre as duas datas e períodos
***********************************************************************************/
  w_atual    date;
  w_inicio   date;
  w_fim      date;
  w_dias     number(10,1);
  w_feriado  number(4,1);
  w_let_ini  date;
  w_let1_fim date;
  w_let2_ini date;
  w_let_fim  date;
  w_publica  varchar2(1);
  w_ano      number;

begin
  w_inicio := to_date(data_inicio,'dd/mm/yyyy');
  w_fim    := to_date(data_fim,'dd/mm/yyyy');
  w_atual  := w_inicio;

  If w_inicio > w_fim or year(w_inicio) <> year(w_fim) Then
     -- Se o período não estiver correto, aponta o erro
     w_dias := -2;
  Else
     w_ano := year(w_inicio);

     -- Verifica se é escola pública ou privada
     select publica into w_publica
       from Cliente a
      where sq_cliente = cliente;

     -- Verifica os inícios e términos de semestre
     If w_publica = 'N' Then
        select dt_ocorrencia into w_let_ini
          from Particular_Calendario         c
               inner join Calendario_Cliente a on (c.sq_particular_calendario = a.sq_particular_calendario)
               inner join Tipo_Data          b on (a.sq_tipo_data = b.sq_tipo_data and year(a.dt_ocorrencia) = w_ano and b.sigla = 'IA')
         where c.sq_cliente = cliente
           and c.sq_particular_calendario = coalesce(calendario,0);

        select dt_ocorrencia into w_let_fim
          from Particular_Calendario         c
               inner join Calendario_Cliente a on (c.sq_particular_calendario = a.sq_particular_calendario)
               inner join Tipo_Data          b on (a.sq_tipo_data = b.sq_tipo_data and year(a.dt_ocorrencia) = w_ano and b.sigla = 'TA')
         where c.sq_cliente = cliente
           and c.sq_particular_calendario = coalesce(calendario,0);

        select dt_ocorrencia into w_let2_ini
          from Particular_Calendario         c
               inner join Calendario_Cliente a on (c.sq_particular_calendario = a.sq_particular_calendario)
               inner join Tipo_Data          b on (a.sq_tipo_data = b.sq_tipo_data and year(a.dt_ocorrencia) = w_ano and b.sigla = 'I2')
         where c.sq_cliente = cliente
           and c.sq_particular_calendario = coalesce(calendario,0);

        select dt_ocorrencia into w_let1_fim
          from Particular_Calendario         c
               inner join Calendario_Cliente a on (c.sq_particular_calendario = a.sq_particular_calendario)
               inner join Tipo_Data          b on (a.sq_tipo_data = b.sq_tipo_data and year(a.dt_ocorrencia) = w_ano and b.sigla = 'T1')
         where c.sq_cliente = cliente
           and c.sq_particular_calendario = coalesce(calendario,0);
     Else
        select dt_ocorrencia into w_let_ini
          from Calendario_Cliente a inner join Tipo_Data b on (a.sq_tipo_data = b.sq_tipo_data and year(a.dt_ocorrencia) = w_ano and b.sigla = 'IA')
         where sq_cliente = 0;

        select dt_ocorrencia into w_let_fim
          from Calendario_Cliente a inner join Tipo_Data b on (a.sq_tipo_data = b.sq_tipo_data and year(a.dt_ocorrencia) = w_ano and b.sigla = 'TA')
         where sq_cliente = 0;

        select dt_ocorrencia into w_let2_ini
          from Calendario_Cliente a inner join Tipo_Data b on (a.sq_tipo_data = b.sq_tipo_data and year(a.dt_ocorrencia) = w_ano and b.sigla = 'I2')
         where sq_cliente = 0;

        select dt_ocorrencia into w_let1_fim
          from Calendario_Cliente a inner join Tipo_Data b on (a.sq_tipo_data = b.sq_tipo_data and year(a.dt_ocorrencia) = w_ano and b.sigla = 'T1')
         where sq_cliente = 0;
     End If;

     If w_let_ini is null or w_let1_fim is null or w_let2_ini is null or w_let_fim is null Then
        -- Se não recuperou alguma das datas, retorna erro.
        Return -2;
     End If;

     w_dias   := 0;
     While w_atual <= w_fim Loop
       -- Recupera somente dentro do ano letivo
       If w_atual between w_let_ini and w_let1_fim or
          w_atual between w_let2_ini and w_let_fim
       Then
          -- Trata as sextas, sábados e domingos
          If to_char(w_atual,'d') not in (1,7) Then
             w_dias := w_dias + 1;
          End If;
       End If;

       -- Incrementa a data atual
       w_atual := w_atual + 1;
     End Loop;

     -- Trata os feriados oficiais
     select count(*) into w_feriado
       from Calendario_Base a
      where a.dt_ocorrencia between w_inicio and w_fim
        and a.dia_letivo = 'N'
        and to_char(a.dt_ocorrencia,'d') not in (1,7)
        and (a.dt_ocorrencia between w_let_ini and w_let1_fim or
             a.dt_ocorrencia between w_let2_ini and w_let_fim
            );

     If w_feriado > 0 Then w_dias := w_dias - w_feriado; End If;

     -- Trata recessos e datas da escola
     If w_publica = 'N' Then
        select count(*) into w_feriado
          from Particular_Calendario         c
               inner join Calendario_Cliente a on (c.sq_particular_calendario = a.sq_particular_calendario)
               inner join Tipo_Data          b on (a.sq_tipo_data             = b.sq_tipo_data)
           where a.sq_cliente           = cliente
           and a.dia_letivo                 = 'N'
           and to_char(a.dt_ocorrencia,'d') not in (1,7)
           and c.sq_particular_calendario   = coalesce(calendario,0)
           and a.dt_ocorrencia  between w_inicio   and w_fim
           and (a.dt_ocorrencia between w_let_ini  and w_let1_fim or
                a.dt_ocorrencia between w_let2_ini and w_let_fim
               );
     Else
        select count(*) into w_feriado
          from Calendario_Cliente   a 
               inner join Tipo_Data b on (a.sq_tipo_data = b.sq_tipo_data)
         where a.sq_cliente             = cliente
           and a.dia_letivo                 = 'N'
           and to_char(a.dt_ocorrencia,'d') not in (1,7)
           and a.dt_ocorrencia  between w_inicio   and w_fim
           and (a.dt_ocorrencia between w_let_ini  and w_let1_fim or
                a.dt_ocorrencia between w_let2_ini and w_let_fim
               );
     End If;

     If w_feriado > 0 Then w_dias := w_dias - w_feriado; End If;

     -- Trata recessos e datas da regional de ensino
     select count(*) into w_feriado
       from Calendario_Cliente    a
            inner join Tipo_Data a1 on (a.sq_tipo_data    = a1.sq_tipo_data)
            inner join Cliente    b on (a.sq_cliente = b.sq_cliente)
            inner join Cliente    c on (b.sq_cliente      = c.sq_cliente_pai)
      where c.sq_cliente                 = cliente
        and a.dia_letivo                 = 'N'
        and to_char(a.dt_ocorrencia,'d') not in (1,7)
        and a.dt_ocorrencia  between w_inicio   and w_fim
        and (a.dt_ocorrencia between w_let_ini  and w_let1_fim or
             a.dt_ocorrencia between w_let2_ini and w_let_fim
            );

     If w_feriado > 0 Then w_dias := w_dias - w_feriado; End If;

     -- Trata recessos e datas da secretaria de educacao
     select count(*) into w_feriado
       from Calendario_Cliente    a
            inner join Tipo_Data a1 on (a.sq_tipo_data    = a1.sq_tipo_data)
            inner join Cliente    b on (a.sq_cliente = b.sq_cliente)
            inner join Cliente    c on (b.sq_cliente      = c.sq_cliente_pai)
            inner join Cliente    d on (c.sq_cliente      = d.sq_cliente_pai)
      where d.sq_cliente                 = cliente
        and a.dia_letivo                 = 'N'
        and to_char(a.dt_ocorrencia,'d') not in (1,7)
        and a.dt_ocorrencia  between w_inicio   and w_fim
        and (a.dt_ocorrencia between w_let_ini  and w_let1_fim or
             a.dt_ocorrencia between w_let2_ini and w_let_fim
            );

     If w_feriado > 0 Then w_dias := w_dias - w_feriado; End If;
  End If;

  Return w_dias;
end DiasLetivos;
/
