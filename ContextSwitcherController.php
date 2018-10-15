<?php

namespace ProspectRelationshipManager\AppBundle\Controller;

use Symfony\Component\HttpFoundation\Request;
use Symfony\Component\HttpFoundation\JsonResponse;
use Sensio\Bundle\FrameworkExtraBundle\Configuration\Route;
use Sensio\Bundle\FrameworkExtraBundle\Configuration\Security;
use ProspectRelationshipManager\AppBundle\Form\Type\ContextSwitcherType;
use JWT\ProspectSdk\Model\GroupsAbstract;
use JWT\ProspectSdk\Model\LeadRecordSortTypeAbstract;

/**
 * @author Scott Richardson ~~~ <scott.richardson@jwt.com>
 *
 * @Route("/userContext")
 */
class ContextSwitcherController extends BaseSdkController
{
    /**
     * @Route("/switcher", name="context_switcher")
     * @Security("is_granted('ROLE_SWITCH_USER_CONTEXT')")
     */
    public function switcherAction(Request $request)
    {
        $switcher = $this->createForm(
                new ContextSwitcherType(
                        $this->getSubStationManagementService(),
                        $this->getUserSessionHelper()
                        )
                );

        if ($request->isMethod('POST')) {
            $switcher->handleRequest($request);
            $newContext = $switcher->getData();

            // *** Update the User Context object in session
            $this->getUserSessionHelper()->updateUserContext(
                $newContext["Group"] == null ? "District" : $newContext["Group"],
                $newContext["SubStation"],
                $newContext["Station"],
                $newContext["District"]
            );

            return new JsonResponse([
                'success' => true,
                'data' => $newContext,
                'href' => $this->generateUrl('homepage'),
            ]);
        }

        $view = array(
            'switcher'      => $switcher->createView(),
            'userContext'   => $this->getUserContext(),
        );

        return $this->render('user_context.html.twig', $view);
    }

  /**
   * Toggles the trusted permissions and assigns the corresponding value to the
   * User Context for trusted access
   *
   * @Route("/toggle/{state}/{startpage}", name="trusted_toggle")
   */
  public function trustedToggleAction(Request $request, $state = "off", $startpage = null)
  {
    $userHelper  = $this->getUserSessionHelper();
    $userContext = $userHelper->getUserContext();
    $userInfo    = $userHelper->getUserInfo();
    $user        = $userHelper->getTokenUser();

    switch ($state) {
      case "off":
        $userInfo->setTrustedState("off");
        $userHelper->resetUserInfo();
        $userHelper->resetUserContext();
        switch ($userContext->getGroup()) {
          case GroupsAbstract::RECRUITER:
            $trustedRoles = [
              'ROLE_RSS_APPROVE',
              'ROLE_RSS_ASSIGN',
              'ROLE_RSS_EXPORT',
              'ROLE_RSS_READ',
              'ROLE_RSS_REASSIGN',
              'ROLE_RSS_TRANSFER',
              'ROLE_RSS_UPDATE'
              ];

            foreach ($trustedRoles as $role) {
              $user->removeRole($role);
            }
            break;
          case GroupsAbstract::SUBSTATION:
            $trustedRoles = [
              'ROLE_RS_APPROVE_PART4',
              'ROLE_RS_ASSIGN',
              'ROLE_RS_EXPORT',
              'ROLE_RS_READ',
              'ROLE_RS_REASSIGN',
              'ROLE_RS_REJECT_INVALID',
              'ROLE_RS_REJECT_PART4',
              'ROLE_RS_TRANSFER',
              'ROLE_RS_UPDATE'
              ];

            foreach ($trustedRoles as $role) {
              $user->removeRole($role);
            }
            break;
          default:
            break;
        }
        break;
      case "on":
        $userInfo->setTrustedState("on");
        //***********************************************
        // Update Context per Trusted Permissions
        //***********************************************
        switch ($userContext->getGroup()) {
          case GroupsAbstract::RECRUITER:
            $userHelper->updateUserContext(GroupsAbstract::SUBSTATION,
                $userContext->getSubStationNumber(),
                $userContext->getStationNumber(),
                $userContext->getDistrictNumber());

            $trustedRoles = [
              'ROLE_RSS_APPROVE',
              'ROLE_RSS_ASSIGN',
              'ROLE_RSS_EXPORT',
              'ROLE_RSS_READ',
              'ROLE_RSS_REASSIGN',
              'ROLE_RSS_TRANSFER',
              'ROLE_RSS_UPDATE'
              ];

            foreach ($trustedRoles as $role) {
              $user->addRole($role);
            }
            break;
          case GroupsAbstract::SUBSTATION:
            $userHelper->updateUserContext(GroupsAbstract::STATION, NULL,
                $userContext->getStationNumber(),
                $userContext->getDistrictNumber());

            $trustedRoles = [
              'ROLE_RS_APPROVE_PART4',
              'ROLE_RS_ASSIGN',
              'ROLE_RS_EXPORT',
              'ROLE_RS_READ',
              'ROLE_RS_REASSIGN',
              'ROLE_RS_REJECT_INVALID',
              'ROLE_RS_REJECT_PART4',
              'ROLE_RS_TRANSFER',
              'ROLE_RS_UPDATE'
              ];
            
            foreach ($trustedRoles as $role) {
              $user->addRole($role);
            }
            break;
          default:
            break;
        }
        //***********************************************
        // Set Trusted Role
        $userHelper->setTrustedPermissions(
            $user, $userContext->getGroup(), $userInfo
        );
        break;
      default:
        break;
    }

    switch($startpage){
      case "rctr":
        return $this->redirectToRoute("prospectList", array(
                    "sortType" => LeadRecordSortTypeAbstract::COMPLETED,
                    "archivedOption" => "NotArchived"
                ));
      case "rss":
        return $this->redirectToRoute("prospectList", array(
                    "sortType" => LeadRecordSortTypeAbstract::UNASSIGNED
                ));
      default:
        return $this->redirect($request->headers->get('referer'));
    }

  }
}
