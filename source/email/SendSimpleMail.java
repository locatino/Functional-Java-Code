import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.FileDataSource;
import javax.mail.*;
import javax.mail.internet.*;
/**
 * ��������java��ʵ��Email�ķ��ͣ����õ���Э��Ϊ��SMTP���˿ں�Ϊ25��
 * ��������Socket����ʵ�֣��򿪿ͻ��˵�Socket���������Ϸ�������
 * �磺Socket sockClient = new Socket("smtp.qq.com",23);
 * ���ʾ���������ӵ���QQ����ķ��������˿ں�Ϊ23
 */
public class SendSimpleMail
{
	/**
	 * ����MIME�ʼ�����
	 */
	private MimeMessage	mimeMsg;
	/**
	 * ר�����������ʼ���Session�Ự
	 */
	private Session		session;
	/**
	 * ��װ�ʼ�����ʱ��һЩ������Ϣ��һ�����Զ���
	 */
	private Properties	props;
	/**
	 * �����˵��û���
	 */
	private String		username;
	/**
	 * �����˵�����
	 */
	private String		password;
	/**
	 * ����ʵ�ָ�����ӵ����
	 */
	private Multipart	mp;
	/**
	 * ���Ͳ�����ʼ��,�еķ���������Ҫ�û���֤������������û�����������г�ʼ��""
	 * 
	 * @param smtp
	 *            SMTP�������ĵ�ַ������Ҫ��QQ���䣬��ôӦΪ��"smtp.qq.com"��163Ϊ��"smtp.163.com"
	 */
	public SendSimpleMail(String smtp)
	{
		username = "";
		password = "";
		// �����ʼ�������
		setSmtpHost(smtp);
		// �����ʼ�
		createMimeMessage();
	}
	/**
	 * ���÷����ʼ�������(JavaMail��ҪProperties������һ��session����
	 * ����Ѱ���ַ���"mail.smtp.host"������ֵ���Ƿ����ʼ�������);
	 * 
	 * @param hostName
	 */
	public void setSmtpHost(String hostName)
	{
		System.out.println("����ϵͳ���ԣ�mail.smtp.host = " + hostName);
		if (props == null)
			props = System.getProperties();
		props.put("mail.smtp.host", hostName);
	}
	/**
	 * (���Session�����JavaMail �е�һ���ʼ�session. ÿһ������
	 * JavaMail��Ӧ�ó���������һ��session���ǿ�����������session�� �����������,
	 * Session������Ҫ֪�����������ʼ���SMTP ��������
	 */
	public boolean createMimeMessage()
	{
		try
		{
			System.out.println("׼����ȡ�ʼ��Ự����");
			// ��props��������������ʼ��session����
			session = Session.getDefaultInstance(props, null);
		}
		catch (Exception e)
		{
			System.err.println("��ȡ�ʼ��Ự����ʱ��������" + e);
			return false;
		}
		System.out.println("׼������MIME�ʼ�����");
		try
		{
			// ��session��������������ʼ���ʼ�����
			mimeMsg = new MimeMessage(session);
			// ���ɸ��������ʵ��
			mp = new MimeMultipart();
		}
		catch (Exception e)
		{
			System.err.println("����MIME�ʼ�����ʧ�ܣ�" + e);
			return false;
		}
		return true;
	}
	/**
	 * ����SMTP�������֤
	 */
	public void setNeedAuth(boolean need)
	{
		System.out.println("����smtp�����֤��mail.smtp.auth = " + need);
		if (props == null)
		{
			props = System.getProperties();
		}
		if (need)
			props.put("mail.smtp.auth", "true");
		else
			props.put("mail.smtp.auth", "false");
	}
	/**
	 * �����û������֤ʱ�������û���������
	 */
	public void setNamePass(String name, String pass)
	{
		System.out.println("����õ��û���������");
		username = name;
		password = pass;
	}
	/**
	 * �����ʼ�����
	 * 
	 * @param mailSubject
	 * @return
	 */
	public boolean setSubject(String mailSubject)
	{
		System.out.println("�����ʼ����⣡");
		try
		{
			mimeMsg.setSubject(mailSubject);
		}
		catch (Exception e)
		{
			System.err.println("�����ʼ����ⷢ������");
			return false;
		}
		return true;
	}
	/**
	 * �����ʼ�����,��������Ϊ�ı���ʽ��HTML�ļ���ʽ�����뷽ʽΪUTF-8
	 * 
	 * @param mailBody
	 * @return
	 */
	public boolean setBody(String mailBody)
	{
		try
		{
			System.out.println("�����ʼ����ʽ");
			BodyPart bp = new MimeBodyPart();
			bp.setContent("<meta http-equiv=Content-Type content=text/html; charset=UTF-8>" + mailBody,
							"text/html;charset=UTF-8");
			// �����������ʼ��ı�
			mp.addBodyPart(bp);
		}
		catch (Exception e)
		{
			System.err.println("�����ʼ�����ʱ��������" + e);
			return false;
		}
		return true;
	}
	/**
	 * ���ӷ��͸���
	 * 
	 * @param filename
	 *            �ʼ������ĵ�ַ��ֻ���Ǳ�����ַ�������������ַ�������׳��쳣 
	 * @return
	 */
	public boolean addFileAffix(String filename)
	{
		System.out.println("�����ʼ�������" + filename);
		try
		{
			BodyPart bp = new MimeBodyPart();
			FileDataSource fileds = new FileDataSource(filename);
			bp.setDataHandler(new DataHandler(fileds));
			// ���͵ĸ���ǰ����һ���û�����ǰ׺
			bp.setFileName(fileds.getName());
			// ��Ӹ���
			mp.addBodyPart(bp);
		}
		catch (Exception e)
		{
			System.err.println("�����ʼ�������" + filename + "��������" + e);
			return false;
		}
		return true;
	}
	/**
	 * ���÷����˵�ַ
	 * 
	 * @param from
	 *            �����˵�ַ
	 * @return
	 */
	public boolean setFrom(String from)
	{
		System.out.println("���÷����ˣ�");
		try
		{
			mimeMsg.setFrom(new InternetAddress(from));
		}
		catch (Exception e)
		{
			return false;
		}
		return true;
	}
	/**
	 * �����ռ��˵�ַ
	 * 
	 * @param to
	 *            �ռ��˵ĵ�ַ
	 * @return
	 */
	public boolean setTo(String[] to)
	{
		System.out.println("����������");
		if (to == null)
			return false;
		try
		{
			InternetAddress[] ia = new InternetAddress[to.length];
			for(int i = 0; i < to.length; i++){
				ia[i] = new InternetAddress(to[0]);
			}
			mimeMsg.setRecipients(javax.mail.Message.RecipientType.TO, ia);
		}
		catch (Exception e)
		{
			return false;
		}
		return true;
	}
	/**
	 * ���͸���
	 * 
	 * @param copyto
	 * @return
	 */
	public boolean setCopyTo(String copyto)
	{
		System.out.println("���͸�����");
		if (copyto == null)
			return false;
		try
		{
			mimeMsg.setRecipients(javax.mail.Message.RecipientType.CC, InternetAddress.parse(copyto));
		}
		catch (Exception e)
		{
			return false;
		}
		return true;
	}
	/**
	 * �����ʼ�
	 * 
	 * @return
	 */
	public boolean sendout()
	{
		try
		{
			mimeMsg.setContent(mp);
			mimeMsg.saveChanges();
			System.out.println("���ڷ����ʼ�....");
			Session mailSession = Session.getInstance(props, null);
			Transport transport = mailSession.getTransport("smtp");
			// �����������ʼ������������������֤
			transport.connect((String) props.get("mail.smtp.host"), username, password);
			// �����ʼ�
			transport.sendMessage(mimeMsg, mimeMsg.getRecipients(javax.mail.Message.RecipientType.TO));
			System.out.println("�����ʼ��ɹ���");
			transport.close();
		}
		catch (Exception e)
		{
			System.err.println("�ʼ�����ʧ�ܣ�" + e.getMessage());
			e.printStackTrace();
			return false;
		}
		return true;
	}
	public static void main(String[] args) throws IOException
	{
		/**
		 *
		 *************����ע��********
		 *
		 *ע��  �ô˳����ʼ���������֧��smtp����  ��2006���Ժ������163�����ǲ�֧�ֵ�
		 *��֪��sina����  sohu���� qq����֧��  ����sina��qq������Ҫ�ֹ����ÿ����˹���	
		 *�����ڲ���ʱ���ʹ������������	
		 *sina�����smtp���÷�������
		 *��¼sina���� ���ε�� ��������--->�ʻ�--->POP/SMTP���� ��������ѡ��ѡ�� Ȼ�󱣴�
		 *
		 *************����ע��********
		 */
		Properties properties = new Properties();
		FileInputStream in = new FileInputStream("mail.properties");
		properties.load(in);
		System.out.println(properties.getProperty("smtpServer"));
//		SendSimpleMail themail = new SendSimpleMail("smtp.qq.com");//������qq����Ϊ����
		SendSimpleMail themail = new SendSimpleMail("113.108.16.44");//������qq����Ϊ����
		String mailbody = new String(properties.getProperty("mailBody").getBytes("ISO8859-1"),"UTF-8");//�ʼ�����
		themail.setNeedAuth(true);
		themail.setSubject("JAVA���ʼ��Ĳ���");//�ʼ�����
		themail.setBody(mailbody);//�ʼ�����
		
		//������Ϣ�����޸�
		themail.setTo(new String[]{"*********@locatino.com"});//�ռ��˵�ַ
		themail.setFrom("*******@qq.com");//�����˵�ַ
		themail.addFileAffix("F:/package.txt");// �����ļ�·��,���磺C:/222.jpg,*ע��"/"��д���� ���û�п��Բ�д
		themail.setNamePass("********@qq.com", "********");//�����˵�ַ������ **��Ϊ��Ӧ��������
		themail.sendout();
	}
}

