   M S B A M L   `   `   `  ����P  LGP MOGen Verification, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null=#clr-namespace:GP_MOGen_VerificationGP_MOGen_Verification  5   	   &     GP_MOGen_Verification.MainWindow   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   5local#clr-namespace:GP_MOGen_Verification     5   	     � Title$  
MainWindow��5   	   $��144.703��6   $��559.368��6-    � 
ResizeMode$ NoResize=�6=    � Icon$ truck_hero_badge.ico��6S   .��5	      � ���	#FFC1C3C9 � RowDefinitions 5
   
   �� 5       � ColumnDefinitions 5   
   �� 5      .+�5   
   �� -   5         ��Name$ Model_Textbox��5      $
��Left=�6)   $��26��6D   $��
132,10,0,0q�6P    �� TextWrapping$
 Wrap=�6d   $	��Top=�6x   $	��172��6�   $��18(�6�   �� 5   
   -   5         ��Name$ 	textBlock��5      $
��Left=�6'   $��17,9,0,0q�6B   $
��Wrap=�6T   $��Item Number:��6h   $	��Top=�6|   $��18(�6�   �� 5   
   -   5         ��Name$ SerializedCheckbox��5      $��Serialized?��6/   $��304,6,145,73q�6E   $��18(�6[   $��Center=�6i   $��Center=�6�   6�   �� 5   
   -   5        $ 
textBlock1��5      $
��Left=�6(   $��	10,41,0,0q�6C   $
��Wrap=�6V   $��Serial Number:��6j   $	��Top=�6�   $��18(�6�   �� 5   
   -   5        $ 
SN_Textbox��5      $
��Left=�6&   $��26��6A   $��
132,41,0,0q�6M   $
 Wrap=�6a   $	��Top=�6u   $	��172��6�   $��18(�6�   {�.  6�   �� 5   
   -   5        	 ��Name$	 VerifyButton��5      $��Verify��6'   $��468,14,10,0q�68   $	��Top=�6M   6e   �� 5   
   -   5        $	 InsertButton��5      $��Insert��6'   $��468,45,10,0q�68   $	��Top=�6M   6e   �� 5   
   -   5        $ 
textBlock2��5      $
��Left=�6(   $��	94,70,0,0q�6C   $
��Wrap=�6V   $��Site:��6j   $	��Top=�6w   $��18(�6�   �� 5   
   -	   5        $	��cmb��5      
 �� DisplayMemberPath$
 SiteID��6    �� SelectedValuePath$ SiteID��69   $��
132,70,0,0q�6T   $��0q�6h   $��0q�6|    � UseLayoutRounding .  J� SnapsToDevicePixels . $	��172��6�   $��26��6�   $	��Top=�6�   $
��Left=�6�   $��16(�6  �� 5   
   -
   5        $ 
textBlock3��5      $
��Left=�6(   $��
309,41,0,0q�6C   $
��Wrap=�6W   $
��Qty:��6k   $	��Top=�6w   $��18(�6�   �� 5   
   -   5        $ 
qtyTextbox��5      $
��Left=�6&   $��26��6A   $��
346,41,0,0q�6M   $
 Wrap=�6a   $	��Top=�6u   $��67��6�   $��18(�6�   �� 5   
   -   5        $	 ReportButton��5      $��Report��6'   $��468,76,10,0q�68   $	��Top=�6M   6e   5      5      